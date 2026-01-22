/*  IP Calculation Functions - Test Cases
    Power Query (M) Implementation Tests
    2025-01-22

    This file contains test cases to validate the IpCalc.pq implementation.
    Import IpCalc.pq first, then run these tests in Power Query.
*/

section IpCalcTests;

//==============================================
//   IPv4 Basic Tests
//==============================================

TestIpIsValid =
[
    Test1 = [
        Name = "Valid IP",
        Input = "192.168.1.1",
        Expected = true,
        Actual = IpCalc.IpIsValid("192.168.1.1"),
        Pass = IpCalc.IpIsValid("192.168.1.1") = true
    ],
    Test2 = [
        Name = "Leading zero (invalid)",
        Input = "192.168.001.1",
        Expected = false,
        Actual = IpCalc.IpIsValid("192.168.001.1"),
        Pass = IpCalc.IpIsValid("192.168.001.1") = false
    ],
    Test3 = [
        Name = "Too many octets",
        Input = "192.168.1.1.1",
        Expected = false,
        Actual = IpCalc.IpIsValid("192.168.1.1.1"),
        Pass = IpCalc.IpIsValid("192.168.1.1.1") = false
    ],
    Test4 = [
        Name = "Too few octets",
        Input = "192.168.1",
        Expected = false,
        Actual = IpCalc.IpIsValid("192.168.1"),
        Pass = IpCalc.IpIsValid("192.168.1") = false
    ]
];

TestIpConversion =
[
    Test1 = [
        Name = "IpStrToNbr",
        Input = "1.2.3.4",
        Expected = 16909060,
        Actual = IpCalc.IpStrToNbr("1.2.3.4"),
        Pass = IpCalc.IpStrToNbr("1.2.3.4") = 16909060
    ],
    Test2 = [
        Name = "IpBinToStr",
        Input = 16909060,
        Expected = "1.2.3.4",
        Actual = IpCalc.IpBinToStr(16909060),
        Pass = IpCalc.IpBinToStr(16909060) = "1.2.3.4"
    ],
    Test3 = [
        Name = "Round-trip conversion",
        Input = "192.168.1.100",
        Expected = "192.168.1.100",
        Actual = IpCalc.IpBinToStr(IpCalc.IpStrToNbr("192.168.1.100")),
        Pass = IpCalc.IpBinToStr(IpCalc.IpStrToNbr("192.168.1.100")) = "192.168.1.100"
    ]
];

TestIpBitwiseOps =
[
    Test1 = [
        Name = "IpAnd",
        Input = ["192.168.1.1", "255.255.255.0"],
        Expected = "192.168.1.0",
        Actual = IpCalc.IpAnd("192.168.1.1", "255.255.255.0"),
        Pass = IpCalc.IpAnd("192.168.1.1", "255.255.255.0") = "192.168.1.0"
    ],
    Test2 = [
        Name = "IpOr",
        Input = ["192.168.1.1", "0.0.0.255"],
        Expected = "192.168.1.255",
        Actual = IpCalc.IpOr("192.168.1.1", "0.0.0.255"),
        Pass = IpCalc.IpOr("192.168.1.1", "0.0.0.255") = "192.168.1.255"
    ],
    Test3 = [
        Name = "IpXor",
        Input = ["192.168.1.1", "0.0.0.255"],
        Expected = "192.168.1.254",
        Actual = IpCalc.IpXor("192.168.1.1", "0.0.0.255"),
        Pass = IpCalc.IpXor("192.168.1.1", "0.0.0.255") = "192.168.1.254"
    ]
];

//==============================================
//   IPv4 Arithmetic Tests
//==============================================

TestIpArithmetic =
[
    Test1 = [
        Name = "IpAdd - same octet",
        Input = ["192.168.1.1", 4],
        Expected = "192.168.1.5",
        Actual = IpCalc.IpAdd("192.168.1.1", 4),
        Pass = IpCalc.IpAdd("192.168.1.1", 4) = "192.168.1.5"
    ],
    Test2 = [
        Name = "IpAdd - carry over",
        Input = ["192.168.1.1", 256],
        Expected = "192.168.2.1",
        Actual = IpCalc.IpAdd("192.168.1.1", 256),
        Pass = IpCalc.IpAdd("192.168.1.1", 256) = "192.168.2.1"
    ],
    Test3 = [
        Name = "IpDiff",
        Input = ["192.168.1.7", "192.168.1.1"],
        Expected = 6,
        Actual = IpCalc.IpDiff("192.168.1.7", "192.168.1.1"),
        Pass = IpCalc.IpDiff("192.168.1.7", "192.168.1.1") = 6
    ]
];

TestIpByteOps =
[
    Test1 = [
        Name = "IpGetByte - first octet",
        Input = ["192.168.1.1", 1],
        Expected = 192,
        Actual = IpCalc.IpGetByte("192.168.1.1", 1),
        Pass = IpCalc.IpGetByte("192.168.1.1", 1) = 192
    ],
    Test2 = [
        Name = "IpGetByte - last octet",
        Input = ["192.168.1.1", 4],
        Expected = 1,
        Actual = IpCalc.IpGetByte("192.168.1.1", 4),
        Pass = IpCalc.IpGetByte("192.168.1.1", 4) = 1
    ],
    Test3 = [
        Name = "IpSetByte",
        Input = ["192.168.1.1", 4, 20],
        Expected = "192.168.1.20",
        Actual = IpCalc.IpSetByte("192.168.1.1", 4, 20),
        Pass = IpCalc.IpSetByte("192.168.1.1", 4, 20) = "192.168.1.20"
    ]
];

//==============================================
//   IPv4 Subnet Tests
//==============================================

TestIpSubnetOps =
[
    Test1 = [
        Name = "IpMask - CIDR notation",
        Input = "192.168.1.1/24",
        Expected = "255.255.255.0",
        Actual = IpCalc.IpMask("192.168.1.1/24"),
        Pass = IpCalc.IpMask("192.168.1.1/24") = "255.255.255.0"
    ],
    Test2 = [
        Name = "IpMask - dotted notation",
        Input = "192.168.1.1 255.255.255.0",
        Expected = "255.255.255.0",
        Actual = IpCalc.IpMask("192.168.1.1 255.255.255.0"),
        Pass = IpCalc.IpMask("192.168.1.1 255.255.255.0") = "255.255.255.0"
    ],
    Test3 = [
        Name = "IpWildMask",
        Input = "192.168.1.1/24",
        Expected = "0.0.0.255",
        Actual = IpCalc.IpWildMask("192.168.1.1/24"),
        Pass = IpCalc.IpWildMask("192.168.1.1/24") = "0.0.0.255"
    ],
    Test4 = [
        Name = "IpInvertMask - mask to wildcard",
        Input = "255.255.255.0",
        Expected = "0.0.0.255",
        Actual = IpCalc.IpInvertMask("255.255.255.0"),
        Pass = IpCalc.IpInvertMask("255.255.255.0") = "0.0.0.255"
    ],
    Test5 = [
        Name = "IpMaskLen",
        Input = "255.255.255.0",
        Expected = 24,
        Actual = IpCalc.IpMaskLen("255.255.255.0"),
        Pass = IpCalc.IpMaskLen("255.255.255.0") = 24
    ],
    Test6 = [
        Name = "IpWithoutMask - CIDR",
        Input = "192.168.1.0/24",
        Expected = "192.168.1.0",
        Actual = IpCalc.IpWithoutMask("192.168.1.0/24"),
        Pass = IpCalc.IpWithoutMask("192.168.1.0/24") = "192.168.1.0"
    ],
    Test7 = [
        Name = "IpSubnetLen - CIDR",
        Input = "192.168.1.1/24",
        Expected = 24,
        Actual = IpCalc.IpSubnetLen("192.168.1.1/24"),
        Pass = IpCalc.IpSubnetLen("192.168.1.1/24") = 24
    ],
    Test8 = [
        Name = "IpSubnetSize",
        Input = "192.168.1.32/29",
        Expected = 8,
        Actual = IpCalc.IpSubnetSize("192.168.1.32/29"),
        Pass = IpCalc.IpSubnetSize("192.168.1.32/29") = 8
    ],
    Test9 = [
        Name = "IpClearHostBits",
        Input = "192.168.1.1/24",
        Expected = "192.168.1.0/24",
        Actual = IpCalc.IpClearHostBits("192.168.1.1/24"),
        Pass = IpCalc.IpClearHostBits("192.168.1.1/24") = "192.168.1.0/24"
    ]
];

TestIpSubnetMatching =
[
    Test1 = [
        Name = "IpIsInSubnet - match",
        Input = ["192.168.1.35", "192.168.1.32/29"],
        Expected = true,
        Actual = IpCalc.IpIsInSubnet("192.168.1.35", "192.168.1.32/29"),
        Pass = IpCalc.IpIsInSubnet("192.168.1.35", "192.168.1.32/29") = true
    ],
    Test2 = [
        Name = "IpIsInSubnet - no match",
        Input = ["192.168.1.41", "192.168.1.32/29"],
        Expected = false,
        Actual = IpCalc.IpIsInSubnet("192.168.1.41", "192.168.1.32/29"),
        Pass = IpCalc.IpIsInSubnet("192.168.1.41", "192.168.1.32/29") = false
    ],
    Test3 = [
        Name = "IpSubnetIsInSubnet - true",
        Input = ["192.168.1.35/30", "192.168.1.32/29"],
        Expected = true,
        Actual = IpCalc.IpSubnetIsInSubnet("192.168.1.35/30", "192.168.1.32/29"),
        Pass = IpCalc.IpSubnetIsInSubnet("192.168.1.35/30", "192.168.1.32/29") = true
    ],
    Test4 = [
        Name = "IpSubnetIsInSubnet - false (wrong subnet)",
        Input = ["192.168.1.41/30", "192.168.1.32/29"],
        Expected = false,
        Actual = IpCalc.IpSubnetIsInSubnet("192.168.1.41/30", "192.168.1.32/29"),
        Pass = IpCalc.IpSubnetIsInSubnet("192.168.1.41/30", "192.168.1.32/29") = false
    ],
    Test5 = [
        Name = "IpIsPrivate - true",
        Input = "192.168.1.35",
        Expected = true,
        Actual = IpCalc.IpIsPrivate("192.168.1.35"),
        Pass = IpCalc.IpIsPrivate("192.168.1.35") = true
    ],
    Test6 = [
        Name = "IpIsPrivate - false",
        Input = "209.85.148.104",
        Expected = false,
        Actual = IpCalc.IpIsPrivate("209.85.148.104"),
        Pass = IpCalc.IpIsPrivate("209.85.148.104") = false
    ]
];

//==============================================
//   IPv4 Array Function Tests
//==============================================

TestIpArrayOps =
[
    Test1 = [
        Name = "IpSortArray",
        Input = {"192.168.1.10", "10.0.0.1", "192.168.1.5"},
        Expected = {"10.0.0.1", "192.168.1.5", "192.168.1.10"},
        Actual = IpCalc.IpSortArray({"192.168.1.10", "10.0.0.1", "192.168.1.5"}),
        Pass = IpCalc.IpSortArray({"192.168.1.10", "10.0.0.1", "192.168.1.5"}){0} = "10.0.0.1"
    ],
    Test2 = [
        Name = "IpSubnetSortArray",
        Input = {"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"},
        Actual = IpCalc.IpSubnetSortArray({"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"}),
        Pass = IpCalc.IpSubnetSortArray({"192.168.1.0/24", "10.0.0.0/8", "192.168.1.32/29"}){0} = "10.0.0.0/8"
    ],
    Test3 = [
        Name = "IpDivideSubnet",
        Input = ["1.2.3.0/24", 2],
        Expected = {"1.2.3.0/26", "1.2.3.64/26", "1.2.3.128/26", "1.2.3.192/26"},
        Actual = IpCalc.IpDivideSubnet("1.2.3.0/24", 2),
        Pass = List.Count(IpCalc.IpDivideSubnet("1.2.3.0/24", 2)) = 4
    ]
];

//==============================================
//   IPv6 Basic Tests
//==============================================

TestIpv6Basic =
[
    Test1 = [
        Name = "Ipv6IsValid - valid",
        Input = "2001:db8::1",
        Expected = true,
        Actual = IpCalc.Ipv6IsValid("2001:db8::1"),
        Pass = IpCalc.Ipv6IsValid("2001:db8::1") = true
    ],
    Test2 = [
        Name = "Ipv6Expand",
        Input = "1:2:3::8",
        Expected = "0001:0002:0003:0000:0000:0000:0000:0008",
        Actual = IpCalc.Ipv6Expand("1:2:3::8"),
        Pass = IpCalc.Ipv6Expand("1:2:3::8") = "0001:0002:0003:0000:0000:0000:0000:0008"
    ],
    Test3 = [
        Name = "Ipv6Compress",
        Input = "0001:0002:0003:0000:0000:0000:0000:0008",
        Expected = "1:2:3::8",
        Actual = IpCalc.Ipv6Compress("0001:0002:0003:0000:0000:0000:0000:0008"),
        Pass = IpCalc.Ipv6Compress("0001:0002:0003:0000:0000:0000:0000:0008") = "1:2:3::8"
    ],
    Test4 = [
        Name = "Ipv6MaskLen",
        Input = "2001:db8:1f89::/48",
        Expected = 48,
        Actual = IpCalc.Ipv6MaskLen("2001:db8:1f89::/48"),
        Pass = IpCalc.Ipv6MaskLen("2001:db8:1f89::/48") = 48
    ],
    Test5 = [
        Name = "Ipv6WithoutMask",
        Input = "2001:db8:1f89::/48",
        Expected = "2001:db8:1f89::",
        Actual = IpCalc.Ipv6WithoutMask("2001:db8:1f89::/48"),
        Pass = IpCalc.Ipv6WithoutMask("2001:db8:1f89::/48") = "2001:db8:1f89::"
    ]
];

TestIpv6Arithmetic =
[
    Test1 = [
        Name = "Ipv6AddInt",
        Input = ["1::2", 16],
        Expected = "1::12",
        Actual = IpCalc.Ipv6AddInt("1::2", 16),
        Pass = IpCalc.Ipv6AddInt("1::2", 16) = "1::12"
    ],
    Test2 = [
        Name = "Ipv6Add",
        Input = ["1:2::", "::3"],
        Expected = "1:2::3",
        Actual = IpCalc.Ipv6Add("1:2::", "::3"),
        Pass = IpCalc.Ipv6Add("1:2::", "::3") = "1:2::3"
    ]
];

TestIpv6Blocks =
[
    Test1 = [
        Name = "Ipv6GetBlock",
        Input = ["2001:db8:1f89:c5a3::ac1f:8001", 2],
        Expected = "0db8",
        Actual = IpCalc.Ipv6GetBlock("2001:db8:1f89:c5a3::ac1f:8001", 2),
        Pass = IpCalc.Ipv6GetBlock("2001:db8:1f89:c5a3::ac1f:8001", 2) = "0db8"
    ],
    Test2 = [
        Name = "Ipv6SetBlock",
        Input = ["2001::", 2, "db8"],
        Expected = "2001:db8::",
        Actual = IpCalc.Ipv6SetBlock("2001::", 2, "db8"),
        Pass = IpCalc.Ipv6SetBlock("2001::", 2, "db8") = "2001:db8::"
    ]
];

TestIpv6Subnet =
[
    Test1 = [
        Name = "Ipv6IsInSubnet - match",
        Input = ["2001:db8:1::ac1f:1", "2001:db8:1::/48"],
        Expected = true,
        Actual = IpCalc.Ipv6IsInSubnet("2001:db8:1::ac1f:1", "2001:db8:1::/48"),
        Pass = IpCalc.Ipv6IsInSubnet("2001:db8:1::ac1f:1", "2001:db8:1::/48") = true
    ],
    Test2 = [
        Name = "Ipv6IsInSubnet - no match",
        Input = ["2001:db8:2::ac1f:1", "2001:db8:1::/48"],
        Expected = false,
        Actual = IpCalc.Ipv6IsInSubnet("2001:db8:2::ac1f:1", "2001:db8:1::/48"),
        Pass = IpCalc.Ipv6IsInSubnet("2001:db8:2::ac1f:1", "2001:db8:1::/48") = false
    ]
];

//==============================================
//   Summary
//==============================================

AllTests = [
    IpIsValid = TestIpIsValid,
    IpConversion = TestIpConversion,
    IpBitwiseOps = TestIpBitwiseOps,
    IpArithmetic = TestIpArithmetic,
    IpByteOps = TestIpByteOps,
    IpSubnetOps = TestIpSubnetOps,
    IpSubnetMatching = TestIpSubnetMatching,
    IpArrayOps = TestIpArrayOps,
    Ipv6Basic = TestIpv6Basic,
    Ipv6Arithmetic = TestIpv6Arithmetic,
    Ipv6Blocks = TestIpv6Blocks,
    Ipv6Subnet = TestIpv6Subnet
];
