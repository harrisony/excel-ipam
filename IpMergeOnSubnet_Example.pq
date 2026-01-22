/*  IpMergeOnSubnet - Example Usage

    This function joins two tables based on subnet containment.
    It's similar to Excel's VLOOKUP or SQL JOIN, but matches IPs to subnets.

    FUNCTION SIGNATURE:
    IpMergeOnSubnet(
        leftTable as table,      // Table with IP addresses
        leftIpColumn as text,    // Column name containing IPs
        rightTable as table,     // Table with subnets
        rightSubnetColumn as text, // Column name containing subnets
        optional joinKind as text  // "Left", "Inner", "Right", or "Full" (default: "Left")
    ) as table

    JOIN TYPES:
    - "Left":  All rows from left table, matching right data where available (default)
    - "Inner": Only rows where IP matches a subnet
    - "Right": All rows from right table, matching left data where available
    - "Full":  All rows from both tables

    BEST MATCH LOGIC:
    When an IP matches multiple subnets, returns the most specific (longest prefix) match.
    Example: 192.168.1.50 matches both 192.168.1.0/24 and 192.168.1.32/27
             → Returns 192.168.1.32/27 (more specific)

    NOTATION SUPPORT:
    Supports both CIDR and dotted mask notation:
    - CIDR: "192.168.1.0/24"
    - Dotted: "192.168.1.0 255.255.255.0"
*/

let
    // Load IpCalc module
    IpCalc = Expression.Evaluate(Text.FromBinary(File.Contents("/Users/harrisony/Labs/excel-ipam/IpCalc.pq")), #shared),

    // Example 1: Basic Usage - Asset Management
    // Match devices to their network segments
    DeviceTable = #table(
        {"IP", "Device", "Owner"},
        {
            {"192.168.1.10", "Server1", "IT Dept"},
            {"192.168.1.50", "Server2", "IT Dept"},
            {"10.0.0.5", "Workstation", "Finance"},
            {"172.16.5.10", "Printer", "Admin"}
        }
    ),

    NetworkTable = #table(
        {"Subnet", "Location", "VLAN"},
        {
            {"192.168.1.0/24", "Data Center", "100"},
            {"10.0.0.0/8", "Office Network", "10"},
            {"172.16.0.0/12", "Remote Sites", "20"}
        }
    ),

    Example1_Result = IpCalc[IpMergeOnSubnet](
        DeviceTable,
        "IP",
        NetworkTable,
        "Subnet",
        "Left"  // All devices, with network info where available
    ),

    // Example 2: Security Audit - Find Devices in Specific Networks
    // Inner join: Only show devices that are in known networks
    Example2_Result = IpCalc[IpMergeOnSubnet](
        DeviceTable,
        "IP",
        NetworkTable,
        "Subnet",
        "Inner"  // Only matched devices
    ),

    // Example 3: Network Planning - Find Unused Networks
    // Right join: Show all networks, see which ones have devices
    Example3_Result = IpCalc[IpMergeOnSubnet](
        DeviceTable,
        "IP",
        NetworkTable,
        "Subnet",
        "Right"  // All networks, with device info where available
    ),

    // Example 4: Best Match - Overlapping Subnets
    // When multiple subnets contain an IP, return the most specific
    OverlappingNetworks = #table(
        {"Subnet", "Description"},
        {
            {"192.168.1.0/24", "Entire Office Floor"},
            {"192.168.1.32/27", "Server Room"},  // More specific
            {"192.168.1.64/28", "Management Network"}  // More specific
        }
    ),

    ServerIPs = #table(
        {"IP", "Server"},
        {
            {"192.168.1.10", "Web Server"},    // Matches 192.168.1.0/24 only
            {"192.168.1.50", "DB Server"},     // Matches 192.168.1.32/27 (best)
            {"192.168.1.70", "File Server"}    // Matches 192.168.1.64/28 (best)
        }
    ),

    Example4_Result = IpCalc[IpMergeOnSubnet](
        ServerIPs,
        "IP",
        OverlappingNetworks,
        "Subnet",
        "Inner"
    ),

    // Example 5: Full Inventory - All Assets and Networks
    // Full join: Complete view of everything
    Example5_Result = IpCalc[IpMergeOnSubnet](
        DeviceTable,
        "IP",
        NetworkTable,
        "Subnet",
        "Full"  // Everything from both tables
    ),

    // Example 6: Dotted Mask Notation
    LegacyNetworks = #table(
        {"Subnet", "Description"},
        {
            {"192.168.1.0 255.255.255.0", "Office LAN"},
            {"10.0.0.0 255.0.0.0", "Corporate WAN"}
        }
    ),

    Example6_Result = IpCalc[IpMergeOnSubnet](
        DeviceTable,
        "IP",
        LegacyNetworks,
        "Subnet",
        "Left"
    ),

    // Select which example to display
    // Change this to see different examples
    Output = Example1_Result

in
    Output
