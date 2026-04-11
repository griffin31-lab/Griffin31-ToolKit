# SPF-Lookup-Validator

## Why this tool?

SPF records allow **up to 10 DNS lookups** (RFC 7208). Every `include`, `a`, `mx`, `ptr`, `exists`, and `redirect` counts as a lookup — and nested includes add up fast. Exceeding 10 causes a **PermError**, and receiving mail servers will treat your emails as unauthenticated.

Most admins add services over time without realizing they've blown past the limit. This tool recursively walks your entire SPF chain and gives you the real count.

**RFC reference:** [RFC 7208 — Sender Policy Framework](https://datatracker.ietf.org/doc/html/rfc7208)

## What it does

- Recursively resolves all SPF includes and redirects
- Counts actual DNS lookups across the full chain
- Validates SPF syntax per RFC 7208
- Warns about deprecated mechanisms (`ptr`), risky qualifiers (`+all`, `?all`), and record length issues
- Accepts a domain name or a raw SPF string as input

## Requirements

- Bash (Linux / macOS)
- `dig` (built-in on macOS, `sudo apt install dnsutils` on Debian/Ubuntu)

## Quick start

```bash
chmod +x spf.sh
./spf.sh
# Enter a domain (e.g. yourdomain.com) or a full SPF record
```

## Files

| File | Purpose |
|------|---------|
| `spf.sh` | Main script |
| `examples.txt` | Sample SPF records to test with |
