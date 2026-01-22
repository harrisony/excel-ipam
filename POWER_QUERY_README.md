# IP Address Calculation Functions - Power Query (M) Implementation

Complete implementation of IP address manipulation functions for Excel Power Query and Power BI.

## Installation

### For Excel Power Query

1. Open Excel
2. Go to **Data** > **Get Data** > **From Other Sources** > **Blank Query**
3. In the Power Query Editor, go to **Home** > **Advanced Editor**
4. Copy the entire contents of `IpCalc.pq` into the editor
5. Click **Done**
6. Rename the query to "IpCalc" in the **Name** field on the right
7. Click **Close & Load**

### For Power BI

1. Open Power BI Desktop
2. Go to **Home** > **Get Data** > **Blank Query**
3. In the Power Query Editor, click **Advanced Editor**
4. Copy the entire contents of `IpCalc.pq` into the editor
5. Click **Done**
6. Rename the query to "IpCalc"
7. Click **Close & Apply**

## Usage Examples

### Basic IP Operations

```m
// Validate IP address
IpCalc.IpIsValid("192.168.1.1")  // returns true
IpCalc.IpIsValid("192.168.001.1")  // returns false (leading zero)

// Convert IP to/from number
IpCalc.IpStrToNbr("1.2.3.4")  // returns 16909060
IpCalc.IpBinToStr(16909060)  // returns "1.2.3.4"

// IP arithmetic
IpCalc.IpAdd("192.168.1.1", 4)  // returns "192.168.1.5"
IpCalc.IpAdd("192.168.1.1", 256)  // returns "192.168.2.1"
IpCalc.IpDiff("192.168.1.7", "192.168.1.1")  // returns 6

// Bitwise operations
IpCalc.IpAnd("192.168.1.1", "255.255.255.0")  // returns "192.168.1.0"
IpCalc.IpOr("192.168.1.1", "0.0.0.255")  // returns "192.168.1.255"
IpCalc.IpXor("192.168.1.1", "0.0.0.255")  // returns "192.168.1.254"
```

### Byte Operations

```m
// Get/Set individual octets
IpCalc.IpGetByte("192.168.1.1", 1)  // returns 192
IpCalc.IpGetByte("192.168.1.1", 4)  // returns 1
IpCalc.IpSetByte("192.168.1.1", 4, 20)  // returns "192.168.1.20"
```

### Subnet Operations

```m
// Both CIDR and dotted mask notations are supported
IpCalc.IpMask("192.168.1.1/24")  // returns "255.255.255.0"
IpCalc.IpMask("192.168.1.1 255.255.255.0")  // returns "255.255.255.0"

IpCalc.IpWildMask("192.168.1.1/24")  // returns "0.0.0.255"
IpCalc.IpInvertMask("255.255.255.0")  // returns "0.0.0.255"

IpCalc.IpMaskLen("255.255.255.0")  // returns 24
IpCalc.IpSubnetLen("192.168.1.1/24")  // returns 24
IpCalc.IpSubnetSize("192.168.1.32/29")  // returns 8

IpCalc.IpWithoutMask("192.168.1.0/24")  // returns "192.168.1.0"
IpCalc.IpClearHostBits("192.168.1.1/24")  // returns "192.168.1.0/24"
```

### Subnet Matching

```m
// Check if IP is in subnet
IpCalc.IpIsInSubnet("192.168.1.35", "192.168.1.32/29")  // returns true
IpCalc.IpIsInSubnet("192.168.1.41", "192.168.1.32/29")  // returns false

// Check if subnet is in subnet
IpCalc.IpSubnetIsInSubnet("192.168.1.35/30", "192.168.1.32/29")  // returns true

// Check for private addresses
IpCalc.IpIsPrivate("192.168.1.35")  // returns true
IpCalc.IpIsPrivate("209.85.148.104")  // returns false
```

### Subnet Lookup

```m
// Match against a list of subnets
let
    subnets = {
        {"192.168.1.0/24", "LAN"},
        {"192.168.1.32/29", "Servers"},
        {"10.0.0.0/8", "Private"}
    }
in
    IpCalc.IpSubnetMatch("192.168.1.35", subnets)  // returns 2 (best match)

// Lookup value from matched subnet
let
    subnets = {
        {"192.168.1.0/24", "LAN"},
        {"10.0.0.0/8", "Private"}
    }
in
    IpCalc.IpSubnetVLookup("192.168.1.35", subnets, 2)  // returns "LAN"
```

### Array Operations

```m
// Sort IP addresses
IpCalc.IpSortArray({"192.168.1.10", "10.0.0.1", "192.168.1.5"})
// returns {"10.0.0.1", "192.168.1.5", "192.168.1.10"}

// Sort subnets
IpCalc.IpSubnetSortArray({"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"})

// Find overlapping subnets
IpCalc.IpFindOverlappingSubnets({"192.168.1.0/24", "192.168.1.32/29", "10.0.0.0/8"})
// returns {"", "192.168.1.0/24", ""}

// Aggregate subnets (remove included, join contiguous, remove duplicates)
IpCalc.IpSubnetAggregateArray({"192.168.1.0/25", "192.168.1.128/25"})
// returns {"192.168.1.0/24"}

// Divide subnet into smaller subnets
IpCalc.IpDivideSubnet("1.2.3.0/24", 2)
// returns {"1.2.3.0/26", "1.2.3.64/26", "1.2.3.128/26", "1.2.3.192/26"}

// Convert IP range to CIDR
IpCalc.IpRangeToCIDR("10.0.0.1", "10.0.0.254")  // returns {"10.0.0.0/24"}

// Subtract subnets
IpCalc.IpSubtractSubnets({"192.168.1.0/24"}, {"192.168.1.0/25"})
// returns {"192.168.1.128/25"}

// Find common subnets
IpCalc.IpCommonSubnets({"192.168.0.0/23"}, {"192.168.1.0/24"})
// returns {"192.168.1.0/24"}
```

### IPv6 Operations

```m
// Validate IPv6 address
IpCalc.Ipv6IsValid("2001:db8::1")  // returns true

// Expand/Compress IPv6
IpCalc.Ipv6Expand("1:2:3::8")
// returns "0001:0002:0003:0000:0000:0000:0000:0008"

IpCalc.Ipv6Compress("0001:0002:0003:0000:0000:0000:0000:0008")
// returns "1:2:3::8"

// IPv6 arithmetic
IpCalc.Ipv6AddInt("1::2", 16)  // returns "1::12"
IpCalc.Ipv6Add("1:2::", "::3")  // returns "1:2::3"

// Block operations
IpCalc.Ipv6GetBlock("2001:db8:1f89:c5a3::ac1f:8001", 2)  // returns "0db8"
IpCalc.Ipv6GetBlockInt("2001:db8:1f89:c5a3::ac1f:8001", 2)  // returns 3512
IpCalc.Ipv6SetBlock("2001::", 2, "db8")  // returns "2001:db8::"
IpCalc.Ipv6SetBlockInt("2001::", 2, 3512)  // returns "2001:db8::"

// IPv4 in IPv6
IpCalc.Ipv6GetIpv4("2001:c0a8:102::", 2)  // returns "192.168.1.2"
IpCalc.Ipv6SetIpv4("2001::", 2, "192.168.1.2")  // returns "2001:c0a8:102::"

// IPv6 subnet operations
IpCalc.Ipv6IsInSubnet("2001:db8:1::ac1f:1", "2001:db8:1::/48")  // returns true
IpCalc.Ipv6SubnetFirstAddress("2001:db8:1:1a0::/59")  // returns "2001:db8:1:1a0::"
IpCalc.Ipv6SubnetLastAddress("2001:db8:1:1a0::/59")
// returns "2001:db8:1:1bf:ffff:ffff:ffff:ffff"

// Sort IPv6 addresses
IpCalc.Ipv6SortArray({"2001:db8::1", "2001:db8::10", "2001:db8::2"})
```

## Using in Power Query Transformations

### Example 1: Add IP Category Column

```m
let
    Source = Excel.CurrentWorkbook(){[Name="IpTable"]}[Content],
    AddCategory = Table.AddColumn(
        Source,
        "Category",
        each if IpCalc.IpIsPrivate([IPAddress]) then "Private" else "Public"
    )
in
    AddCategory
```

### Example 2: Filter IPs in Subnet

```m
let
    Source = Excel.CurrentWorkbook(){[Name="IpTable"]}[Content],
    FilteredRows = Table.SelectRows(
        Source,
        each IpCalc.IpIsInSubnet([IPAddress], "192.168.1.0/24")
    )
in
    FilteredRows
```

### Example 3: Sort IPs Properly

```m
let
    Source = Excel.CurrentWorkbook(){[Name="IpTable"]}[Content],
    IpList = Source[IPAddress],
    Sorted = IpCalc.IpSortArray(IpList),
    ToTable = Table.FromList(Sorted, Splitter.SplitByNothing(), {"IPAddress"})
in
    ToTable
```

### Example 4: Aggregate Subnet Allocations

```m
let
    Source = Excel.CurrentWorkbook(){[Name="Subnets"]}[Content],
    SubnetList = Source[Subnet],
    Aggregated = IpCalc.IpSubnetAggregateArray(SubnetList),
    ToTable = Table.FromList(Aggregated, Splitter.SplitByNothing(), {"Aggregated Subnet"})
in
    ToTable
```

## Function Reference

See [IMPLEMENTATION_STATUS.md](IMPLEMENTATION_STATUS.md) for a complete list of all 46+ functions with descriptions and examples.

## Key Features

### Dual Notation Support

All subnet functions support both notations:
- **CIDR notation**: `192.168.1.0/24`
- **Dotted mask notation**: `192.168.1.0 255.255.255.0`

The original notation is preserved in function outputs.

### Functional & Immutable

All functions are pure and immutable, following Power Query M best practices:
- No side effects
- Deterministic results
- Thread-safe operations

### Performance Optimized

- Uses Power Query native functions where possible
- Efficient list operations with `List.Accumulate` and `List.Transform`
- Optimized sorting and comparison algorithms

## Testing

A comprehensive test suite is provided in `IpCalc_Tests.pq`. Import this file to validate the implementation.

## Compatibility

- **Excel 2016+** with Power Query
- **Excel for Microsoft 365**
- **Power BI Desktop**
- **Azure Data Factory** (Data Flows)

## Language Resources

- [Power Query M Formula Language](https://learn.microsoft.com/en-us/powerquery-m/)
- [Power Query M Primer](https://bengribaudo.com/power-query-m-primer)

## License

GNU General Public License v3.0 or later

## Related Implementations

This Power Query implementation provides identical functionality to:
- **VBA** (`ipcalc_module.bas`) - For Excel/LibreOffice macros
- **JavaScript** (`ip-calc.js`) - For Google Sheets

All three implementations share the same function names (accounting for naming conventions) and produce identical results.
