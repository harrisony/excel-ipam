# IpMergeOnSubnet - Quick Reference

## Syntax

```powerquery
IpMergeOnSubnet(leftTable, leftIpColumn, rightTable, rightSubnetColumn, [joinKind])
```

## Quick Examples

### Show all devices with network info
```powerquery
IpMergeOnSubnet(Devices, "IP", Networks, "Subnet", "Left")
// or omit "Left" (it's the default)
IpMergeOnSubnet(Devices, "IP", Networks, "Subnet")
```

### Show only matched devices
```powerquery
IpMergeOnSubnet(Devices, "IP", Networks, "Subnet", "Inner")
```

### Show all networks with device counts
```powerquery
IpMergeOnSubnet(Devices, "IP", Networks, "Subnet", "Right")
```

### Complete view of everything
```powerquery
IpMergeOnSubnet(Devices, "IP", Networks, "Subnet", "Full")
```

## Join Types Comparison

| Join Type | Returns | Use When |
|-----------|---------|----------|
| **Left** (default) | All left rows + matching right data | "Show all my devices" |
| **Inner** | Only matched rows | "Show only known devices" |
| **Right** | All right rows + matching left data | "Show all networks" |
| **Full** | All rows from both tables | "Show everything" |

## Common Patterns

### Pattern 1: Enrich device list with network info
```powerquery
let
    Source = Excel.CurrentWorkbook(){[Name="Devices"]}[Content],
    Networks = Excel.CurrentWorkbook(){[Name="Networks"]}[Content],
    Result = IpMergeOnSubnet(Source, "IP Address", Networks, "Subnet", "Left")
in
    Result
```

### Pattern 2: Find devices in production network
```powerquery
let
    Source = Excel.CurrentWorkbook(){[Name="Devices"]}[Content],
    ProdNets = #table({"Subnet"}, {{"10.1.0.0/16"}, {"10.2.0.0/16"}}),
    Result = IpMergeOnSubnet(Source, "IP", ProdNets, "Subnet", "Inner")
in
    Result
```

### Pattern 3: Identify unused subnets
```powerquery
let
    Devices = Excel.CurrentWorkbook(){[Name="Devices"]}[Content],
    AllSubnets = Excel.CurrentWorkbook(){[Name="Subnets"]}[Content],
    Merged = IpMergeOnSubnet(Devices, "IP", AllSubnets, "Subnet", "Right"),
    Unused = Table.SelectRows(Merged, each [IP] = null)
in
    Unused
```

## Tips

✅ **DO:**
- Use "Left" join to keep all your data
- Use "Inner" join to filter to known networks
- Use "Right" join to analyze network utilization
- Test with small datasets first

❌ **DON'T:**
- Forget the join type defaults to "Left"
- Assume all IPs will match (use Left join to keep unmatched)
- Forget that best match returns most specific subnet

## Subnet Formats

Both formats work:
- `192.168.1.0/24` (CIDR notation)
- `192.168.1.0 255.255.255.0` (dotted mask)

## Best Match Example

If an IP matches multiple subnets, returns the most specific:

```
IP: 192.168.1.50

Subnets:
  192.168.1.0/24    ← broader network (256 IPs)
  192.168.1.32/27   ← more specific (32 IPs) ✓ MATCHED

Result: Returns 192.168.1.32/27
```

## Troubleshooting

### No matches found
- Check IP column name is exact (case-sensitive)
- Check subnet column name is exact
- Verify subnet format: `192.168.1.0/24` or `192.168.1.0 255.255.255.0`
- Try Inner join to see if any rows match

### Wrong subnet matched
- Check for overlapping subnets
- Function always returns most specific match
- Use aggregation functions to debug matches

### Error: "Invalid joinKind"
- Must be exactly: "Left", "Inner", "Right", or "Full"
- Check capitalization (case-sensitive)

## Need More Help?

- See **IpMergeOnSubnet_README.md** for full documentation
- See **IpMergeOnSubnet_Example.pq** for 6 working examples
- See **IpMergeOnSubnet_Test.pq** for test cases
