# IpMergeOnSubnet Function

## Overview

`IpMergeOnSubnet` is a Power Query function that performs table joins based on subnet containment. It matches IP addresses in one table with subnet ranges in another table, similar to SQL JOIN or Excel VLOOKUP but with subnet-based matching logic.

## Function Signature

```powerquery
IpMergeOnSubnet(
    leftTable as table,           // Table containing IP addresses
    leftIpColumn as text,          // Name of column with IP addresses
    rightTable as table,           // Table containing subnets
    rightSubnetColumn as text,     // Name of column with subnets
    optional joinKind as text      // Join type: "Left", "Inner", "Right", "Full"
) as table
```

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `leftTable` | table | The table containing IP addresses to match |
| `leftIpColumn` | text | Name of the column in leftTable containing IP addresses |
| `rightTable` | table | The table containing subnet ranges |
| `rightSubnetColumn` | text | Name of the column in rightTable containing subnet definitions |
| `joinKind` | text (optional) | Join type - defaults to "Left" if not specified |

## Join Types

### Left Join (Default)
- Returns all rows from `leftTable`
- Adds matching data from `rightTable` where available
- Non-matching rows have NULL values for `rightTable` columns
- **Use case:** "Show all devices with their network information if available"

### Inner Join
- Returns only rows where the IP address matches a subnet
- Excludes rows with no match
- **Use case:** "Show only devices that are in known networks"

### Right Join
- Returns all rows from `rightTable`
- Adds matching data from `leftTable` where available
- Networks with no matching IPs have NULL values for `leftTable` columns
- **Use case:** "Show all networks and which devices are in them"

### Full Join
- Returns all rows from both tables
- Combines results of Left and Right joins
- **Use case:** "Complete inventory of devices and networks"

## Best Match Logic

When an IP address falls within multiple subnets, the function returns the **most specific match** (longest prefix):

**Example:**
- IP: `192.168.1.50`
- Subnets:
  - `192.168.1.0/24` (covers 192.168.1.0 - 192.168.1.255)
  - `192.168.1.32/27` (covers 192.168.1.32 - 192.168.1.63) ✓ More specific

Result: Matches `192.168.1.32/27` because it's more specific (longer prefix)

## Notation Support

The function supports both subnet notation formats:

### CIDR Notation
```
192.168.1.0/24
10.0.0.0/8
172.16.0.0/12
```

### Dotted Mask Notation
```
192.168.1.0 255.255.255.0
10.0.0.0 255.0.0.0
172.16.0.0 255.240.0.0
```

## Examples

### Example 1: Basic Asset Management

Match devices to their network segments:

```powerquery
let
    DeviceTable = #table(
        {"IP", "Device", "Owner"},
        {
            {"192.168.1.10", "Server1", "IT"},
            {"10.0.0.5", "Workstation", "Finance"}
        }
    ),

    NetworkTable = #table(
        {"Subnet", "Location"},
        {
            {"192.168.1.0/24", "Data Center"},
            {"10.0.0.0/8", "Office"}
        }
    ),

    Result = IpMergeOnSubnet(
        DeviceTable,
        "IP",
        NetworkTable,
        "Subnet",
        "Left"
    )
in
    Result
```

**Output:**

| IP | Device | Owner | Subnet | Location |
|----|--------|-------|--------|----------|
| 192.168.1.10 | Server1 | IT | 192.168.1.0/24 | Data Center |
| 10.0.0.5 | Workstation | Finance | 10.0.0.0/8 | Office |

### Example 2: Security Audit

Find devices in specific networks only:

```powerquery
IpMergeOnSubnet(DeviceTable, "IP", NetworkTable, "Subnet", "Inner")
```

Only returns devices that match a known network (excludes unknowns).

### Example 3: Network Utilization

See which networks have devices:

```powerquery
IpMergeOnSubnet(DeviceTable, "IP", NetworkTable, "Subnet", "Right")
```

Shows all networks, indicating which ones have devices and which are empty.

### Example 4: Overlapping Subnets

Demonstrate best match with overlapping subnets:

```powerquery
let
    IPs = #table({"IP"}, {{"192.168.1.50"}}),

    Subnets = #table(
        {"Subnet", "Description"},
        {
            {"192.168.1.0/24", "Entire Floor"},
            {"192.168.1.32/27", "Server Room"}  // More specific
        }
    ),

    Result = IpMergeOnSubnet(IPs, "IP", Subnets, "Subnet", "Inner")
in
    Result
```

**Output:**

| IP | Subnet | Description |
|----|--------|-------------|
| 192.168.1.50 | 192.168.1.32/27 | Server Room |

Note: Matches the more specific subnet (192.168.1.32/27) even though the IP is also in 192.168.1.0/24.

## Error Handling

### NULL and Empty Values
- NULL or empty IP addresses do not match any subnet
- NULL or empty subnets are skipped during matching
- Invalid IP or subnet formats are ignored (no error thrown)

### Invalid Join Kind
If an invalid `joinKind` is specified, the function returns an error:
```
Invalid joinKind. Must be one of: Left, Inner, Right, Full
```

### Column Name Conflicts
If both tables have columns with the same name (other than the join columns), Power Query will automatically suffix column names with `.1`, `.2`, etc.

## Performance Considerations

- **Time Complexity:** O(n × m) where n = rows in leftTable, m = rows in rightTable
- **Best Practices:**
  - For large datasets, consider filtering tables before joining
  - Index tables if your data source supports it
  - Use Inner join when possible to reduce result size

## Use Cases

1. **Asset Management:** Match devices to their network segments
2. **Security Auditing:** Identify devices in specific networks
3. **Network Planning:** Find unused or underutilized subnets
4. **IP Address Management (IPAM):** Track IP allocations across subnets
5. **Compliance Reporting:** Verify devices are in correct network zones
6. **Network Troubleshooting:** Identify which subnet a problematic IP belongs to

## Testing

Test files are provided:

- **IpMergeOnSubnet_Test.pq** - Comprehensive test suite with 9 test cases
- **IpMergeOnSubnet_Example.pq** - 6 practical examples with sample data

To run tests:
1. Open Power Query Editor
2. Create new Blank Query
3. Open Advanced Editor
4. Paste test/example file contents
5. Click Done

## Implementation Notes

- Function is marked as `shared` to be callable from queries
- Uses existing internal functions: `_ParseIpNet`, `_IpAnd`
- Follows Power Query naming convention: PascalCase
- Compatible with both Excel Power Query and Power BI

## Compatibility with Other Implementations

This function is **new** and does not yet exist in the VBA or JavaScript implementations. Consider adding equivalent functions to maintain consistency across all three implementations:

- **VBA:** `IpMergeOnSubnet()` function
- **JavaScript:** `ipMergeOnSubnet()` function

## Version History

- **v1.0** (2025-01-22) - Initial implementation
  - Supports all four join types (Left, Inner, Right, Full)
  - Best match logic for overlapping subnets
  - CIDR and dotted mask notation support
  - NULL/empty value handling

## License

This function is part of the excel-ipam project and is licensed under the GNU General Public License v3.0 or later.

Copyright 2010-2025 Thomas Rohmer-Kretz
