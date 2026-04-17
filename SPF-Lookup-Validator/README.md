# SPF-Lookup-Validator

> **RFC 7208-compliant SPF chain analysis** — recursively counts DNS lookups against the 10-lookup limit.

<sub>[← Back to Griffin31 ToolKit](../)</sub>

---

## What you get

- **Real DNS lookup count** across the full include chain
- **Warnings** for deprecated mechanisms (`ptr`), risky qualifiers (`+all`, `?all`), record length issues
- **Full include tree** showing nested SPF resolution
- **Syntax validation** per RFC 7208
- **Accepts domain OR raw SPF string** as input

## Quick start

```bash
chmod +x spf.sh
./spf.sh
# Enter a domain (e.g. yourdomain.com) or paste a full SPF record
```

## Why this tool?

SPF records allow **up to 10 DNS lookups** (RFC 7208). Every `include`, `a`, `mx`, `ptr`, `exists`, and `redirect` counts as a lookup — and nested includes add up fast. Exceeding 10 causes **PermError**, and receiving mail servers treat your email as unauthenticated.

Most admins add services over time without realizing they've blown past the limit. This tool walks the full chain and gives you the real count.

Reference: [RFC 7208 — Sender Policy Framework](https://datatracker.ietf.org/doc/html/rfc7208)

## Requirements

- Bash (macOS / Linux)
- `dig` — built-in on macOS; `sudo apt install dnsutils` on Debian/Ubuntu

## Files

| File | Purpose |
|------|---------|
| `spf.sh` | Main script |
| `examples.txt` | Sample SPF records to test with |

## Related tools

Any email-security gap surfaced here often pairs with DMARC + DKIM issues — those are worth checking in MXToolbox or a dedicated email-auth tool after you've cleaned up SPF.
