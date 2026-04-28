#!/usr/bin/env python3
"""
SMTP Email Verifier · standalone Python script for Free Mail Merge.

Verifies email deliverability without sending, by:
  1. Checking the domain has MX records (DNS).
  2. Connecting to the highest-priority MX server.
  3. Issuing HELO + MAIL FROM + RCPT TO (RFC 5321 conversation).
  4. Reading the SMTP response code:
        250 → email accepted (likely deliverable)
        550 → recipient rejected (does not exist)
        4xx → temporary failure (greylisting, retry later)
        anything else → uncertain
  5. Detecting catch-all domains (servers that accept anything).

Catch-all detection: probes a clearly-fake address (e.g. random16chars@domain).
If that also returns 250, the domain is catch-all and per-address verification
is impossible. Marked as "catch_all".

Usage as CLI (single email):
    python3 smtp_verifier.py john@acme.com

Usage as CLI (CSV batch — one email per line, or first column of CSV):
    python3 smtp_verifier.py --csv emails.csv > verified.csv

Usage as HTTP server (for Apps Script integration):
    python3 smtp_verifier.py --serve [--port 8080]
    # then: GET /verify?email=john@acme.com
    # response: {"email":"...","status":"verified|not_found|catch_all|error","mx":"...","code":250,"message":"..."}

No external dependencies (stdlib only). Works on macOS, Linux, any Python 3.7+.

Notes:
  - Some ISPs block outbound port 25. Run from a server with port 25 open.
  - Gmail/Outlook MX always return 250 (catch-all). These will be marked catch_all.
  - Use a real-looking sender (HELO domain + MAIL FROM) to avoid being blocked.
  - Rate-limit yourself: < 10 verifications/min/domain to avoid blocklists.

Repo: https://github.com/diegotorreslopez81/free-mail-merge
License: MIT
"""
from __future__ import annotations

import argparse
import csv
import json
import random
import re
import smtplib
import socket
import string
import struct
import sys
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse

# ---------------- Config ----------------

DEFAULT_HELO = "verifier.local"
DEFAULT_FROM = "verify@" + DEFAULT_HELO
SMTP_TIMEOUT = 8  # seconds per SMTP step
DNS_TIMEOUT = 4

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

# Known consumer catch-all providers — these always 250 RCPT TO, so the probe
# is meaningless. Detected only on consumer free accounts (gmail.com,
# yahoo.com etc.), NOT on Google Workspace / Microsoft 365 corporate domains
# routed via aspmx / protection.outlook.com — those do answer RCPT TO realistically.
KNOWN_CONSUMER_CATCHALL_DOMAINS = {
    "gmail.com", "googlemail.com",
    "yahoo.com", "yahoo.es", "yahoo.co.uk", "ymail.com",
    "outlook.com", "hotmail.com", "live.com", "msn.com",
    "icloud.com", "me.com", "mac.com",
    "aol.com",
}


# ---------------- DNS MX lookup (no dependencies) ----------------

def _dns_query(domain: str, qtype: int = 15, server: str = "8.8.8.8") -> bytes:
    """Send a raw DNS query for `domain` (qtype 15 = MX). Returns response bytes."""
    txid = random.randint(0, 0xFFFF)
    flags = 0x0100  # standard query, RD
    header = struct.pack(">HHHHHH", txid, flags, 1, 0, 0, 0)
    qname = b""
    for label in domain.encode("idna").split(b"."):
        qname += bytes([len(label)]) + label
    qname += b"\x00"
    question = qname + struct.pack(">HH", qtype, 1)  # IN class
    pkt = header + question

    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.settimeout(DNS_TIMEOUT)
    s.sendto(pkt, (server, 53))
    try:
        data, _ = s.recvfrom(4096)
    finally:
        s.close()
    return data


def _parse_dns_name(packet: bytes, offset: int):
    parts = []
    while True:
        length = packet[offset]
        if length == 0:
            offset += 1
            break
        if (length & 0xC0) == 0xC0:  # pointer
            ptr = ((length & 0x3F) << 8) | packet[offset + 1]
            sub, _ = _parse_dns_name(packet, ptr)
            parts.append(sub)
            offset += 2
            return ".".join(parts), offset
        offset += 1
        parts.append(packet[offset:offset + length].decode("ascii"))
        offset += length
    return ".".join(parts), offset


def _parse_mx_response(data: bytes):
    if len(data) < 12:
        return []
    qcount = struct.unpack(">H", data[4:6])[0]
    acount = struct.unpack(">H", data[6:8])[0]
    if acount == 0:
        return []
    offset = 12
    # skip questions
    for _ in range(qcount):
        _, offset = _parse_dns_name(data, offset)
        offset += 4
    answers = []
    for _ in range(acount):
        _, offset = _parse_dns_name(data, offset)
        rtype, _, _, rdlen = struct.unpack(">HHIH", data[offset:offset + 10])
        offset += 10
        if rtype == 15:  # MX
            pref = struct.unpack(">H", data[offset:offset + 2])[0]
            mx, _ = _parse_dns_name(data, offset + 2)
            answers.append((pref, mx))
        offset += rdlen
    answers.sort(key=lambda a: a[0])
    return [mx for _, mx in answers]


def get_mx_records(domain: str):
    """Return MX hostnames for domain, sorted by priority. Empty if no MX."""
    try:
        data = _dns_query(domain)
        return _parse_mx_response(data)
    except (socket.timeout, OSError):
        return []


# ---------------- SMTP RCPT TO probe ----------------

def smtp_probe(mx: str, mail_from: str, rcpt_to: str, helo: str = DEFAULT_HELO):
    """
    Connect to mx:25, do HELO/MAIL/RCPT, return (code, message).
    Does NOT send DATA. Closes connection cleanly with QUIT.
    """
    try:
        with smtplib.SMTP(mx, 25, timeout=SMTP_TIMEOUT, local_hostname=helo) as s:
            s.helo(helo)
            s.mail(mail_from)
            code, msg = s.rcpt(rcpt_to)
            try:
                s.quit()
            except Exception:
                pass
            return code, (msg or b"").decode(errors="replace")
    except smtplib.SMTPServerDisconnected:
        return 0, "disconnected"
    except smtplib.SMTPResponseException as e:
        return e.smtp_code, str(e.smtp_error)
    except (socket.timeout, OSError, ConnectionRefusedError) as e:
        return 0, str(e)
    except Exception as e:
        return 0, f"{type(e).__name__}: {e}"


def random_local_part(n: int = 16) -> str:
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=n))


def verify_email(email: str, mail_from: str = DEFAULT_FROM, helo: str = DEFAULT_HELO,
                 known_catchall_check: bool = True) -> dict:
    """
    Verify a single email. Returns dict:
      {email, status: verified|not_found|catch_all|invalid_format|no_mx|temp_fail|error,
       mx: <hostname>, code: int, message: str, catch_all_probe_code: int|None}
    """
    out = {"email": email, "status": "error", "mx": None, "code": 0, "message": ""}
    if not EMAIL_RE.match(email or ""):
        out["status"] = "invalid_format"
        out["message"] = "Email failed regex"
        return out
    domain = email.split("@", 1)[1].lower()

    # Consumer free providers always 250 RCPT TO regardless — skip probe.
    if known_catchall_check and domain in KNOWN_CONSUMER_CATCHALL_DOMAINS:
        out["status"] = "catch_all"
        out["message"] = f"Consumer provider {domain} (cannot be verified per-address)"
        return out

    mx_list = get_mx_records(domain)
    if not mx_list:
        out["status"] = "no_mx"
        out["message"] = "No MX records"
        return out
    mx = mx_list[0]
    out["mx"] = mx

    # Probe real RCPT TO
    code, msg = smtp_probe(mx, mail_from, email, helo)
    out["code"] = code
    out["message"] = msg

    if code == 250:
        # Need to confirm not catch-all: probe random@domain
        fake = random_local_part() + "@" + domain
        fake_code, _ = smtp_probe(mx, mail_from, fake, helo)
        out["catch_all_probe_code"] = fake_code
        if fake_code == 250:
            out["status"] = "catch_all"
        else:
            out["status"] = "verified"
    elif code in (550, 553, 554):
        # Distinguish "user does not exist" from "your IP is blocked by RBL/Spamhaus".
        # If the rejection is reputation-based, every email on this domain will return
        # 550 — meaningless. Mark as error so caller knows to run from a clean IP.
        msg_lower = msg.lower()
        ip_blocked_markers = ("spamhaus", "blocklist", "blocked using",
                              "5.7.1 service unavailable", "blacklisted",
                              "denied due to your reputation",
                              "barracuda", "client host")
        if any(m in msg_lower for m in ip_blocked_markers):
            out["status"] = "error"
            out["message"] = "Sender IP blocked by RBL — run smtp_verifier.py from a clean IP/VPS"
        else:
            out["status"] = "not_found"
    elif 400 <= code < 500:
        out["status"] = "temp_fail"
    else:
        out["status"] = "error"
    return out


# ---------------- CLI ----------------

def cli_single(email: str):
    res = verify_email(email)
    print(json.dumps(res, indent=2))


def cli_csv(path: str, email_col: str = "email", out_path: str | None = None):
    rows_out = []
    with open(path, newline="") as f:
        reader = csv.DictReader(f)
        if email_col not in (reader.fieldnames or []):
            # fallback: first column
            email_col = (reader.fieldnames or ["email"])[0]
        for r in reader:
            email = (r.get(email_col) or "").strip()
            if not email:
                rows_out.append({**r, "_verifier_status": "skipped", "_verifier_mx": "", "_verifier_code": ""})
                continue
            res = verify_email(email)
            print(f"  · {email[:35]:35} → {res['status']}", file=sys.stderr)
            rows_out.append({
                **r,
                "_verifier_status": res["status"],
                "_verifier_mx": res.get("mx") or "",
                "_verifier_code": str(res.get("code") or ""),
            })
            time.sleep(0.5)  # rate-limit
    fields = list(rows_out[0].keys()) if rows_out else (reader.fieldnames or [])
    if out_path:
        with open(out_path, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(rows_out)
        print(f"Wrote {len(rows_out)} rows to {out_path}", file=sys.stderr)
    else:
        w = csv.DictWriter(sys.stdout, fieldnames=fields)
        w.writeheader()
        w.writerows(rows_out)


# ---------------- HTTP server (for Apps Script integration) ----------------

class VerifierHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # quieter
        sys.stderr.write(f"[{self.log_date_time_string()}] {fmt % args}\n")

    def _json(self, status: int, payload: dict):
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path != "/verify":
            self._json(404, {"error": "use /verify?email=..."})
            return
        qs = parse_qs(u.query)
        email = (qs.get("email") or [""])[0].strip()
        if not email:
            self._json(400, {"error": "missing email"})
            return
        res = verify_email(email)
        self._json(200, res)


def serve(port: int = 8080):
    addr = ("0.0.0.0", port)
    print(f"smtp_verifier serving on http://0.0.0.0:{port}/verify?email=...", file=sys.stderr)
    HTTPServer(addr, VerifierHandler).serve_forever()


# ---------------- Entry point ----------------

def main():
    p = argparse.ArgumentParser(description="SMTP email verifier (RCPT TO probe + catch-all detection)")
    p.add_argument("email", nargs="?", help="email to verify (skip if --csv or --serve)")
    p.add_argument("--csv", help="batch verify a CSV (looks for 'email' column or first col)")
    p.add_argument("--out", help="output CSV path (with --csv); defaults to stdout")
    p.add_argument("--serve", action="store_true", help="run HTTP API on --port (default 8080)")
    p.add_argument("--port", type=int, default=8080)
    args = p.parse_args()

    if args.serve:
        serve(args.port)
    elif args.csv:
        cli_csv(args.csv, out_path=args.out)
    elif args.email:
        cli_single(args.email)
    else:
        p.print_help()
        sys.exit(2)


if __name__ == "__main__":
    main()
