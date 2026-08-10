# JavaScript/VBA/Power Query parity notes

This note records the parity decisions for the current IPv4 Power Query
foundation. `ip-calc.js` is the behavioral cross-check where a corresponding
API or helper exists; `ipcalc_module.bas` remains the inventory source for
VBA functions that JavaScript does not expose directly.

| Power Query | JavaScript reference | Parity decision |
| --- | --- | --- |
| `IpStrToBin` | `ipStrToNbr` → `_ipToNbr` | Preserve base-256 accumulation for valid IPv4 addresses. M uses an explicit four-octet, ASCII-decimal, 0..255 contract instead of JavaScript numeric coercion. |
| `IpBinToStr` | `ipNbrToStr` → `_ipFromNbr` | Preserve four-octet formatting for whole numbers in 0..4294967295. M rejects null, fractional, negative, and out-of-range values explicitly. |
| `IpParse` | No direct public counterpart | Preserve the VBA right-to-left fragment operation, returning `[Byte, Remainder]` because M has no `ByRef` mutation. |
| `IpBuild` | No direct public counterpart; related to `_ipFromNbr` | Preserve low-byte/carry behavior, returning `[Ip, Carry]` as an immutable M record. |
| `IpMaskLen` | `ipMaskLen` | Return 0..32 only for canonical contiguous masks. This deliberately rejects non-contiguous masks instead of JavaScript's last-one-bit interpretation. |
| `IpSubnetParse` | Related `IpNet` parsing | Preserve the VBA parser contract: return address text and prefix length without silently normalizing host bits. Network normalization belongs to a separate operation. |

The M contracts intentionally reject malformed input and null explicitly. These
are contract decisions, not claims that the permissive JavaScript behavior is
equivalent. Future functions must add their JavaScript counterpart, helper
algorithm, and any deliberate divergence to this table before implementation
is marked complete.
