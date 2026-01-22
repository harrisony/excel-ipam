/*  Test file for IpMergeOnSubnet function

    To use this test:
    1. Open Power Query Editor in Excel or Power BI
    2. Create a new Blank Query
    3. Open Advanced Editor
    4. Paste this code
    5. Click Done

    Expected results are documented below each test case.
*/

let
    // Load the IpCalc module (adjust path as needed)
    IpCalc = Expression.Evaluate(Text.FromBinary(File.Contents("/Users/harrisony/Labs/excel-ipam/IpCalc.pq")), #shared),

    // Test Data - IP Table
    IpTable = #table(
        {"IP", "Device"},
        {
            {"192.168.1.10", "Server1"},
            {"192.168.1.50", "Server2"},
            {"10.0.0.5", "Workstation"},
            {"8.8.8.8", "DNS"},
            {"172.16.5.10", "Printer"}
        }
    ),

    // Test Data - Subnet Table (includes overlapping subnets to test best match)
    SubnetTable = #table(
        {"Subnet", "Location", "Type"},
        {
            {"192.168.1.0/24", "Office Floor 1", "LAN"},
            {"192.168.1.32/27", "Office Server Room", "LAN"},  // Overlapping - more specific
            {"10.0.0.0/8", "Corporate Network", "Private"},
            {"172.16.0.0/12", "Remote Sites", "Private"}
        }
    ),

    // Test 1: Inner Join
    // Expected: 4 rows (all except 8.8.8.8 which doesn't match any subnet)
    // 192.168.1.50 should match "192.168.1.32/27" (more specific) not "192.168.1.0/24"
    Test1_InnerJoin = IpCalc[IpMergeOnSubnet](IpTable, "IP", SubnetTable, "Subnet", "Inner"),
    Test1_RowCount = Table.RowCount(Test1_InnerJoin),
    Test1_Expected = 4,
    Test1_Pass = Test1_RowCount = Test1_Expected,

    // Verify Server2 matches the more specific subnet
    Test1_Server2 = Table.SelectRows(Test1_InnerJoin, each [Device] = "Server2"),
    Test1_Server2_Location = Table.Column(Test1_Server2, "Location"){0},
    Test1_BestMatch_Pass = Test1_Server2_Location = "Office Server Room",

    // Test 2: Left Join
    // Expected: 5 rows (all from left table, 8.8.8.8 with null subnet info)
    Test2_LeftJoin = IpCalc[IpMergeOnSubnet](IpTable, "IP", SubnetTable, "Subnet", "Left"),
    Test2_RowCount = Table.RowCount(Test2_LeftJoin),
    Test2_Expected = 5,
    Test2_Pass = Test2_RowCount = Test2_Expected,

    // Verify DNS row has null location
    Test2_DNS = Table.SelectRows(Test2_LeftJoin, each [Device] = "DNS"),
    Test2_DNS_Location = Table.Column(Test2_DNS, "Location"){0},
    Test2_Null_Pass = Test2_DNS_Location = null,

    // Test 3: Right Join
    // Expected: 4 rows (all from right table, some may have multiple left matches)
    Test3_RightJoin = IpCalc[IpMergeOnSubnet](IpTable, "IP", SubnetTable, "Subnet", "Right"),
    Test3_RowCount = Table.RowCount(Test3_RightJoin),
    Test3_HasAllSubnets = Table.RowCount(
        Table.SelectRows(Test3_RightJoin, each [Subnet] <> null)
    ) = 4,
    Test3_Pass = Test3_HasAllSubnets,

    // Test 4: Full Join
    // Expected: At least 5 rows (all IPs + any unmatched subnets)
    Test4_FullJoin = IpCalc[IpMergeOnSubnet](IpTable, "IP", SubnetTable, "Subnet", "Full"),
    Test4_RowCount = Table.RowCount(Test4_FullJoin),
    Test4_Pass = Test4_RowCount >= 5,

    // Test 5: Null/Empty handling
    IpTableWithNulls = Table.InsertRows(IpTable, 0, {
        [IP = null, Device = "Unknown1"],
        [IP = "", Device = "Unknown2"]
    }),
    Test5_NullHandling = IpCalc[IpMergeOnSubnet](IpTableWithNulls, "IP", SubnetTable, "Subnet", "Left"),
    Test5_Pass = Table.RowCount(Test5_NullHandling) = 7,

    // Test 6: Dotted mask notation support
    SubnetTableDotted = #table(
        {"Subnet", "Location"},
        {
            {"192.168.1.0 255.255.255.0", "Office"}
        }
    ),
    Test6_DottedMask = IpCalc[IpMergeOnSubnet](
        #table({"IP"}, {{"192.168.1.10"}}),
        "IP",
        SubnetTableDotted,
        "Subnet",
        "Inner"
    ),
    Test6_Pass = Table.RowCount(Test6_DottedMask) = 1,

    // Test 7: Default join kind (should be Left)
    Test7_DefaultJoin = IpCalc[IpMergeOnSubnet](IpTable, "IP", SubnetTable, "Subnet"),
    Test7_Pass = Table.RowCount(Test7_DefaultJoin) = 5,

    // Summary Results
    TestResults = #table(
        {"Test", "Description", "Result", "Details"},
        {
            {"Test 1", "Inner Join", Test1_Pass, "Expected " & Number.ToText(Test1_Expected) & " rows, got " & Number.ToText(Test1_RowCount)},
            {"Test 1b", "Best Match Logic", Test1_BestMatch_Pass, "Server2 matched to: " & Test1_Server2_Location},
            {"Test 2", "Left Join", Test2_Pass, "Expected " & Number.ToText(Test2_Expected) & " rows, got " & Number.ToText(Test2_RowCount)},
            {"Test 2b", "Null for Unmatched", Test2_Null_Pass, "DNS Location is null: " & (if Test2_DNS_Location = null then "true" else "false")},
            {"Test 3", "Right Join", Test3_Pass, "All subnets present: " & (if Test3_HasAllSubnets then "true" else "false")},
            {"Test 4", "Full Join", Test4_Pass, "Expected >= 5 rows, got " & Number.ToText(Test4_RowCount)},
            {"Test 5", "Null Handling", Test5_Pass, "Expected 7 rows with nulls"},
            {"Test 6", "Dotted Mask", Test6_Pass, "Dotted mask notation works"},
            {"Test 7", "Default Join", Test7_Pass, "Default is Left join"}
        }
    ),

    // Show results - change this to see different outputs
    Output = TestResults  // Change to Test1_InnerJoin, Test2_LeftJoin, etc. to see actual data
in
    Output
