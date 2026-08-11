let
    /* Private validation helpers shared by IPv4 functions. */
    IpIsAsciiDigits = (value as text) as logical =>
        value <> "" and Text.Select(value, {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}) = value,

    /* Creates a validation failure function while keeping each caller's public Reason explicit. */
    IpMakeFailure = (reason as text) as function =>
        (message as text, detail as record) as none =>
            error Error.Record(reason, message, detail),

    /*
        Converts a dotted-decimal IPv4 address to its base-256 integer form.

        The VBA reference accumulates one value for each dot-delimited segment:

            result = result * 256 + octet

        This M version makes the input contract explicit. It accepts exactly four
        ASCII decimal octets, each in the inclusive range 0..255. Leading zeroes
        are accepted because they do not change the numeric value.

        The JavaScript counterpart is ipStrToNbr, implemented through _ipToNbr;
        both use the same base-256 accumulation for valid IPv4 input. M rejects
        malformed or out-of-range input explicitly instead of inheriting
        JavaScript's numeric coercion behavior.

        Example:
            IpStrToBinDoc("1.2.3.4") = 16909060
    */
    IpStrToBin = (ip as nullable text) as number =>
        let
            fail = IpMakeFailure("IpStrToBin.InvalidInput"),

            parseOctet = (octet as text, position as number) as number =>
                let
                    containsOnlyDigits = IpIsAsciiDigits(octet),
                    parsed =
                        if octet = "" then
                            fail(
                                "IPv4 octets cannot be empty.",
                                [Input = ip, Component = "octet", Position = position, Value = octet, Expected = "non-empty ASCII decimal octet"]
                            )
                        else if containsOnlyDigits then
                            Number.FromText(octet, "en-US")
                        else
                            null,
                    isValid =
                        if not containsOnlyDigits then
                            false
                        else
                            parsed >= 0
                                and parsed <= 255
                                and Number.RoundDown(parsed) = parsed
                in
                    if isValid then
                        parsed
                    else
                        fail(
                            "IPv4 octet must be an ASCII decimal integer from 0 through 255.",
                            [
                                Input = ip,
                                Component = "octet",
                                Position = position,
                                Value = octet,
                                Expected = "0..255"
                            ]
                        ),

            octets =
                if ip = null then
                    fail(
                        "IPv4 address cannot be null.",
                        [Input = ip, Component = "address", Expected = "text"]
                    )
                else
                    Text.Split(ip, "."),
            octetCount = List.Count(octets),
            parsedOctets =
                if octetCount <> 4 then
                    fail(
                        "IPv4 address must contain exactly four dot-delimited octets.",
                        [
                            Input = ip,
                            Component = "address",
                            Count = octetCount,
                            Expected = 4
                        ]
                    )
                else
                    List.Transform(
                        List.Positions(octets),
                        (position as number) => parseOctet(octets{position}, position + 1)
                    )
        in
            List.Accumulate(
                parsedOctets,
                0,
                (state as number, octet as number) => state * 256 + octet
            ),

    IpStrToBinType =
        type function (
            ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
            ])
        ) as number meta [
            Documentation.Name = "IpStrToBin",
            Documentation.LongDescription = "Returns the IPv4 address as a number in the range 0 through 4294967295. The input must contain exactly four ASCII decimal octets in the range 0 through 255.",
            Documentation.Examples = {
                [
                    Description = "Convert 1.2.3.4 to its integer representation.",
                    Code = "IpStrToBin(\"1.2.3.4\")",
                    Result = "16909060"
                ],
                [
                    Description = "Convert the first IPv4 address.",
                    Code = "IpStrToBin(\"0.0.0.0\")",
                    Result = "0"
                ],
                [
                    Description = "Convert the last IPv4 address.",
                    Code = "IpStrToBin(\"255.255.255.255\")",
                    Result = "4294967295"
                ]
            }
        ],

    IpStrToBinDoc = Value.ReplaceType(IpStrToBin, IpStrToBinType),

    /*
        Converts an IPv4 integer to dotted-decimal text.

        The VBA reference extracts four base-256 digits from right to left,
        then prepends each digit to the result. The validated M contract is
        the unsigned IPv4 integer range 0..4294967295.

        The JavaScript counterpart is ipNbrToStr, implemented through _ipFromNbr;
        this function preserves its four-octet output for the validated IPv4
        range while making the range and whole-number requirements explicit.

        Example:
            IpBinToStrDoc(16909060) = "1.2.3.4"
    */
    IpBinToStr = (ip as nullable number) as text =>
        let
            fail = IpMakeFailure("IpBinToStr.InvalidInput"),

            checkedIp =
                if ip = null then
                    fail(
                        "IPv4 integer cannot be null.",
                        [Input = ip, Component = "address", Expected = "number"]
                    )
                else if ip < 0 or ip > 4294967295 then
                    fail(
                        "IPv4 integer must be in the range 0 through 4294967295.",
                        [
                            Input = ip,
                            Component = "address",
                            Expected = "0..4294967295"
                        ]
                    )
                else if Number.RoundDown(ip) <> ip then
                    fail(
                        "IPv4 integer must be a whole number.",
                        [
                            Input = ip,
                            Component = "address",
                            Expected = "whole number"
                        ]
                    )
                else
                    ip,

            leastSignificantFirst =
                List.Accumulate(
                    {1..4},
                    [Value = checkedIp, Octets = {}],
                    (state as record, _ as number) as record =>
                        let
                            quotient = Number.IntegerDivide(state[Value], 256),
                            remainder = Number.Mod(state[Value], 256)
                        in
                            [
                                Value = quotient,
                                Octets = state[Octets] & {remainder}
                            ]
                )[Octets],
            octets = List.Reverse(leastSignificantFirst)
        in
            Text.Combine(
                List.Transform(
                    octets,
                    (octet as number) => Number.ToText(octet, "0", "en-US")
                ),
                "."
            ),

    IpBinToStrType =
        type function (
           ip as (type nullable number meta [
                Documentation.FieldCaption = "IPv4 integer",
                Documentation.FieldDescription = "An integer in the range 0 through 4294967295."
           ])
        ) as text meta [
            Documentation.Name = "IpBinToStr",
            Documentation.LongDescription = "Returns a four-octet IPv4 address for an integer in the range 0 through 4294967295.",
            Documentation.Examples = {
                [
                    Description = "Convert 16909060 to dotted-decimal IPv4 text.",
                    Code = "IpBinToStr(16909060)",
                    Result = "1.2.3.4"
                ],
                [
                    Description = "Convert the first IPv4 address.",
                    Code = "IpBinToStr(0)",
                    Result = "0.0.0.0"
                ],
                [
                    Description = "Convert the last IPv4 address.",
                    Code = "IpBinToStr(4294967295)",
                    Result = "255.255.255.255"
                ]
            }
        ],

    IpBinToStrDoc = Value.ReplaceType(IpBinToStr, IpBinToStrType),

    /*
        Tests whether text is exactly a canonical dotted-decimal IPv4 address.

        The VBA implementation converts the address to an integer, formats
        it back to dotted text, and compares that canonical result with the
        original input. The JavaScript counterpart applies the same exact
        comparison and also requires four dot-delimited components.

        M uses the strict IpStrToBin parser and catches its validation error so
        this predicate returns false for null, malformed, out-of-range, or
        non-canonical text instead of allowing an error to escape. Leading
        zeroes and surrounding spaces therefore return false.

        Examples:
            IpIsValidDoc("192.168.1.1") = true
            IpIsValidDoc("192.168.001.1") = false
            IpIsValidDoc("192.168.1") = false
    */
    IpIsValid = (ip as nullable text) as logical =>
        let
            parsed = try IpStrToBin(ip)
        in
            if parsed[HasError] then
                false
            else
                let
                    canonical = try IpBinToStr(parsed[Value])
                in
                    if canonical[HasError] then
                        false
                    else
                        canonical[Value] = ip,

    IpIsValidType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 text",
                Documentation.FieldDescription = "Text that may contain a canonical dotted-decimal IPv4 address."
           ])
        ) as logical meta [
            Documentation.Name = "IpIsValid",
            Documentation.LongDescription = "Returns true only when the input is exactly four ASCII decimal IPv4 octets in canonical dotted-decimal form. Null, malformed, out-of-range, leading-zero, and whitespace-containing inputs return false rather than raising an error.",
            Documentation.Examples = {
                [
                    Description = "Accept a canonical IPv4 address.",
                    Code = "IpIsValid(\"192.168.1.1\")",
                    Result = "true"
                ],
                [
                    Description = "Reject leading zeroes.",
                    Code = "IpIsValid(\"192.168.001.1\")",
                    Result = "false"
                ],
                [
                    Description = "Reject an address with fewer than four octets.",
                    Code = "IpIsValid(\"192.168.1\")",
                    Result = "false"
                ]
            }
        ],

    IpIsValidDoc = Value.ReplaceType(IpIsValid, IpIsValidType),

    /*
        Tests whether an IPv4 address is in an RFC 1918 private range.

        The VBA and JavaScript references test the three private ranges
        10.0.0.0/8, 172.16.0.0/12, and 192.168.0.0/16. M parses the search
        address once, represents each fixed range as a numeric network and
        block size, and uses List.Transform plus List.AnyTrue for the
        membership test. This avoids invoking the subnet parser three times
        for one address while keeping public values as dotted text.

        Null and malformed addresses raise the structured validation error
        from IpStrToBin. Leading-zero octets are accepted for the numeric
        membership test, matching the permissive numeric reference helpers;
        IpIsValid remains the separate canonical-text predicate.

        Examples:
            IpIsPrivateDoc("192.168.1.35") = true
            IpIsPrivateDoc("172.16.0.1") = true
            IpIsPrivateDoc("209.85.148.104") = false
    */
    IpIsPrivate = (ip as nullable text) as logical =>
        let
            addressValue = IpStrToBin(ip),
            privateRanges = {
                [Network = 167772160, BlockSize = 16777216],
                [Network = 2886729728, BlockSize = 1048576],
                [Network = 3232235520, BlockSize = 65536]
            },
            isInPrivateRange = (privateRange as record) as logical =>
                addressValue >= privateRange[Network]
                    and addressValue < privateRange[Network] + privateRange[BlockSize]
        in
            List.AnyTrue(List.Transform(privateRanges, isInPrivateRange)),

    IpIsPrivateType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ])
        ) as logical meta [
            Documentation.Name = "IpIsPrivate",
            Documentation.LongDescription = "Returns true for addresses in 10.0.0.0/8, 172.16.0.0/12, or 192.168.0.0/16. Null and malformed inputs raise structured validation errors; leading-zero octets are accepted for numeric membership testing.",
            Documentation.Examples = {
                [
                    Description = "An address in the 192.168.0.0/16 private range.",
                    Code = "IpIsPrivate(\"192.168.1.35\")",
                    Result = "true"
                ],
                [
                    Description = "An address in the 172.16.0.0/12 private range.",
                    Code = "IpIsPrivate(\"172.16.0.1\")",
                    Result = "true"
                ],
                [
                    Description = "A public address outside all three private ranges.",
                    Code = "IpIsPrivate(\"209.85.148.104\")",
                    Result = "false"
                ]
            }
        ],

    IpIsPrivateDoc = Value.ReplaceType(IpIsPrivate, IpIsPrivateType),

    /*
        Parses the rightmost byte from an IPv4 address fragment.

        The VBA reference removes the rightmost dot-delimited component from
        its ByRef input and returns that component as an integer. M values are
        immutable, so this version returns both values in a record. The
        remainder is an empty text value when the input has no dot.

        Unlike VBA's permissive Val conversion, the M contract requires the
        extracted component to be a non-empty ASCII decimal integer from 0
        through 255. Earlier components remain in Remainder for a subsequent
        parse, matching the reference's right-to-left iteration model.

        Examples:
            IpParseDoc("192.168.1.32")
                = [Byte = 32, Remainder = "192.168.1"]
            IpParseDoc("32")
                = [Byte = 32, Remainder = ""]
    */
    IpParse = (ip as nullable text) as record =>
        let
            fail = IpMakeFailure("IpParse.InvalidInput"),

            checkedIp =
                if ip = null then
                    fail(
                        "IPv4 address fragment cannot be null.",
                        [Input = ip, Component = "address fragment", Expected = "text"]
                    )
                else
                    ip,
            separatorPosition = Text.PositionOf(checkedIp, ".", Occurrence.Last),
            byteText =
                if separatorPosition < 0 then
                    checkedIp
                else
                    Text.Range(checkedIp, separatorPosition + 1),
            remainder =
                if separatorPosition < 0 then
                    ""
                else
                    Text.Start(checkedIp, separatorPosition),
            containsOnlyDigits = IpIsAsciiDigits(byteText),
            byteValue =
                if containsOnlyDigits then
                    Number.FromText(byteText, "en-US")
                else
                    null,
            isValidByte =
                if not containsOnlyDigits then
                    false
                else
                    byteValue >= 0
                        and byteValue <= 255
                        and Number.RoundDown(byteValue) = byteValue
        in
            if isValidByte then
                [Byte = byteValue, Remainder = remainder]
            else
                fail(
                    "IPv4 byte must be an ASCII decimal integer from 0 through 255.",
                    [
                        Input = ip,
                        Component = "byte",
                        Value = byteText,
                        Expected = "0..255"
                    ]
                ),

    IpParseType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address fragment",
                Documentation.FieldDescription = "An IPv4 address or partial address whose rightmost byte is parsed."
           ])
        ) as [Byte = number, Remainder = text] meta [
            Documentation.Name = "IpParse",
            Documentation.LongDescription = "Returns a record containing Byte and Remainder fields. The input is processed from right to left; Remainder is the text before the final dot, or an empty text value when no dot is present. The extracted byte must be an ASCII decimal integer from 0 through 255.",
            Documentation.Examples = {
                [
                    Description = "Parse the rightmost byte and retain the preceding address fragment.",
                    Code = "IpParse(\"192.168.1.32\")",
                    Result = "[Byte = 32, Remainder = \"192.168.1\"]"
                ],
                [
                    Description = "Parse a single byte with no preceding address fragment.",
                    Code = "IpParse(\"32\")",
                    Result = "[Byte = 32, Remainder = \"\"]"
                ],
                [
                    Description = "Parse the largest valid IPv4 byte.",
                    Code = "IpParse(\"255\")",
                    Result = "[Byte = 255, Remainder = \"\"]"
                ]
            }
        ],

    IpParseDoc = Value.ReplaceType(IpParse, IpParseType),

    /*
        Splits a route expression into subnet text and next-hop text.

        The VBA helper returns the subnet through its function result and the
        next hop through a ByRef argument. M values are immutable, so the
        equivalent contract is a record with Subnet and NextHop fields.

        For CIDR routes, the first space starts the next hop. For dotted-mask
        routes, the first space separates the address from its mask and the
        second space starts the next hop. The NextHop field deliberately keeps
        that delimiter-leading suffix exactly as the VBA Mid(route, sp)
        operation does. This helper only splits text; route consumers validate
        the returned Subnet value separately.

        Supported examples:
            IpParseRouteDoc("10.0.0.0/24 1.2.3.4")
                = [Subnet = "10.0.0.0/24", NextHop = " 1.2.3.4"]
            IpParseRouteDoc("10.0.0.0 255.255.255.0 next hop")
                = [Subnet = "10.0.0.0 255.255.255.0", NextHop = " next hop"]
            IpParseRouteDoc("10.0.0.1 gw")
                = [Subnet = "10.0.0.1", NextHop = " gw"]
            IpParseRouteDoc("10.0.0.0 ")
                = [Subnet = "10.0.0.0", NextHop = ""]
    */
    IpParseRoute = (route as nullable text) as [Subnet = text, NextHop = text] =>
        let
            fail = IpMakeFailure("IpParseRoute.InvalidInput"),

            checkedRoute =
                if route = null then
                    fail(
                        "Route text cannot be null.",
                        [Input = route, Component = "route", Expected = "route text"]
                    )
                else
                    route,
            trimmedRoute = Text.Trim(checkedRoute),
            firstSpacePosition = Text.PositionOf(trimmedRoute, " "),
            firstToken = if firstSpacePosition < 0 then trimmedRoute else Text.Start(trimmedRoute, firstSpacePosition),
            slashPosition = Text.PositionOf(trimmedRoute, "/"),
            slashInFirstToken = Text.PositionOf(firstToken, "/") >= 0,
            afterFirstDelimiter = if firstSpacePosition < 0 then "" else Text.Range(trimmedRoute, firstSpacePosition),
            maskLeadingWhitespace = Text.Length(afterFirstDelimiter) - Text.Length(Text.TrimStart(afterFirstDelimiter)),
            textAfterFirstSpace = Text.TrimStart(afterFirstDelimiter),
            secondSpaceRelativePosition = Text.PositionOf(textAfterFirstSpace, " "),
            secondToken = if secondSpaceRelativePosition < 0 then textAfterFirstSpace else Text.Start(textAfterFirstSpace, secondSpaceRelativePosition),
            parsedDottedMask = try IpMaskLen(secondToken),
            isDottedMaskRoute = secondToken <> "" and not parsedDottedMask[HasError],
            splitPosition =
                if slashPosition >= 0 and not slashInFirstToken then
                    fail(
                        "CIDR prefix must be attached to the route address; mixed or slash-after-space notation is invalid.",
                        [Input = route, Component = "route", Expected = "address/prefix [next hop] or address mask [next hop]"]
                    )
                else if firstSpacePosition < 0 then
                    -1
                else if slashInFirstToken or not isDottedMaskRoute then
                    firstSpacePosition
                else if secondSpaceRelativePosition < 0 then
                    -1
                else
                    firstSpacePosition + maskLeadingWhitespace + Text.Length(secondToken),
            subnet = if splitPosition < 0 then trimmedRoute else Text.Start(trimmedRoute, splitPosition),
            nextHop = if splitPosition < 0 then "" else Text.Range(trimmedRoute, splitPosition)
        in
            [Subnet = subnet, NextHop = nextHop],

    IpParseRouteType =
        type function (
           route as (type nullable text meta [
                Documentation.FieldCaption = "Route text",
                Documentation.FieldDescription = "A route containing a CIDR or dotted-mask subnet and an optional next hop."
           ])
        ) as [Subnet = text, NextHop = text] meta [
            Documentation.Name = "IpParseRoute",
            Documentation.LongDescription = "Returns a [Subnet, NextHop] record. CIDR and unmasked host routes split before an optional next hop; dotted-mask route text keeps the address and validated mask together. Leading and trailing route whitespace is trimmed, NextHop retains its delimiter-leading suffix, and a separated CIDR prefix is rejected.",
            Documentation.Examples = {
                [
                    Description = "Split a CIDR route and next hop.",
                    Code = "IpParseRoute(\"10.0.0.0/24 1.2.3.4\")",
                    Result = "[Subnet = \"10.0.0.0/24\", NextHop = \" 1.2.3.4\"]"
                ],
                [
                    Description = "Keep a dotted-mask subnet together before the next hop.",
                    Code = "IpParseRoute(\"10.0.0.0 255.255.255.0 next hop\")",
                    Result = "[Subnet = \"10.0.0.0 255.255.255.0\", NextHop = \" next hop\"]"
                ],
                [
                    Description = "Split an unmasked host route from its next hop.",
                    Code = "IpParseRoute(\"10.0.0.1 gw\")",
                    Result = "[Subnet = \"10.0.0.1\", NextHop = \" gw\"]"
                ]
            }
        ],

    IpParseRouteDoc = Value.ReplaceType(IpParseRoute, IpParseRouteType),

    /*
        Finds containing subnets for each row of an IPv4 subnet list.

        The VBA and JavaScript references preserve one output row per input
        row. A row contains an overlapping subnet when another row contains
        it; the first such row in input order is returned. Rows with no
        containing subnet receive an empty text value. Equal normalized
        subnets therefore count as overlaps when they occur in different
        positions, matching the reference predicate's inclusive prefix test.

        The Power Query contract is a list of subnet text values and the
        result is a same-length list of text values. Null or empty input items,
        non-text values, and malformed subnet expressions raise structured
        validation errors because silently dropping a row would destroy the
        required positional alignment. Returned matches retain the original
        subnet text from the containing row.

        Examples:
            IpFindOverlappingSubnetsDoc({"10.0.0.0/8", "10.1.0.0/16", "192.0.2.0/24"})
                = {"", "10.0.0.0/8", ""}
            IpFindOverlappingSubnetsDoc({"10.0.0.0/24", "10.0.0.0/24"})
                = {"10.0.0.0/24", "10.0.0.0/24"}
            IpFindOverlappingSubnetsDoc({"10.0.0.0/8", "192.0.2.0/24"})
                = {"", ""}
    */
    IpFindOverlappingSubnets = (subnets as nullable list) as list =>
        let
            fail = IpMakeFailure("IpFindOverlappingSubnets.InvalidInput"),

            checkedSubnets =
                if subnets = null then
                    fail(
                        "Subnet list cannot be null.",
                        [Input = subnets, Component = "subnets", Expected = "list of subnet text"]
                    )
                else
                    subnets,
            subnetRecords =
                List.Buffer(
                    List.Transform(
                        List.Positions(checkedSubnets),
                        (position as number) as record =>
                            let
                                rawSubnet = checkedSubnets{position},
                                checkedSubnet =
                                    if not Value.Is(rawSubnet, type text) then
                                        fail(
                                            "Subnet list items must be text.",
                                            [Input = rawSubnet, Component = "subnet", Row = position + 1, Expected = "text"]
                                        )
                                    else if rawSubnet = "" then
                                        fail(
                                            "Subnet list items cannot be empty.",
                                            [Input = rawSubnet, Component = "subnet", Row = position + 1, Expected = "IPv4 subnet text"]
                                        )
                                    else
                                        rawSubnet,
                                parsedSubnetAttempt = try IpSubnetParseWithAddressValue(checkedSubnet, "IpFindOverlappingSubnets.InvalidInput"),
                                parsedSubnet =
                                    if parsedSubnetAttempt[HasError] then
                                        fail(
                                            "Subnet list item is not a valid IPv4 subnet.",
                                            [Input = checkedSubnet, Component = "subnet", Row = position + 1, Cause = parsedSubnetAttempt[Error]]
                                        )
                                    else
                                        parsedSubnetAttempt[Value],
                                facts = parsedSubnet[Facts]
                            in
                                [
                                    Original = checkedSubnet,
                                    NetworkKey = facts[NetworkKey],
                                    PrefixLength = facts[PrefixLength],
                                    BlockSize = facts[BlockSize]
                                ]
                    )
                ),
            contains = (inner as record, outer as record) as logical =>
                inner[PrefixLength] >= outer[PrefixLength]
                    and inner[NetworkKey] >= outer[NetworkKey]
                    and inner[NetworkKey] < outer[NetworkKey] + outer[BlockSize],
            findContainingSubnet = (position as number) as text =>
                let
                    candidate = subnetRecords{position},
                    containingPosition =
                        List.PositionOf(
                            List.Positions(subnetRecords),
                            true,
                            Occurrence.First,
                            (otherPosition as number, _ as logical) as logical =>
                                otherPosition <> position
                                    and contains(candidate, subnetRecords{otherPosition})
                        )
                in
                    if containingPosition < 0 then
                        ""
                    else
                        subnetRecords{containingPosition}[Original],
            result =
                List.Transform(
                    List.Positions(subnetRecords),
                    (position as number) as text => findContainingSubnet(position)
                )
        in
            result,

    IpFindOverlappingSubnetsType =
        type function (
           subnets as (type nullable list meta [
                Documentation.FieldCaption = "IPv4 subnet list",
                Documentation.FieldDescription = "An ordered list of IPv4 subnet text values. The result preserves this list's row alignment."
           ])
        ) as list meta [
            Documentation.Name = "IpFindOverlappingSubnets",
            Documentation.LongDescription = "Scans an ordered Power Query list of IPv4 subnet text and returns a same-length text list. For each row, the first other row whose subnet contains it is returned using the original text; rows with no containing subnet return an empty text value. Duplicate normalized subnets count as overlaps. Null or empty lists/items, non-text values, and malformed subnets raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Preserve row alignment while reporting a containing subnet.",
                    Code = "IpFindOverlappingSubnets({\"10.0.0.0/8\", \"10.1.0.0/16\", \"192.0.2.0/24\"})",
                    Result = "{\"\", \"10.0.0.0/8\", \"\"}"
                ],
                [
                    Description = "Treat duplicate normalized subnets as overlaps.",
                    Code = "IpFindOverlappingSubnets({\"10.0.0.0/24\", \"10.0.0.0/24\"})",
                    Result = "{\"10.0.0.0/24\", \"10.0.0.0/24\"}"
                ],
                [
                    Description = "Return empty text for rows with no overlap.",
                    Code = "IpFindOverlappingSubnets({\"10.0.0.0/8\", \"192.0.2.0/24\"})",
                    Result = "{\"\", \"\"}"
                ]
            }
        ],

    IpFindOverlappingSubnetsDoc = Value.ReplaceType(IpFindOverlappingSubnets, IpFindOverlappingSubnetsType),

    /*
        Prepends the low byte of an IPv4 calculation to an address fragment.

        VBA returns the remaining base-256 carry while mutating its ByRef
        address argument. M values are immutable, so this version returns a
        record containing both results. The low eight bits are represented by
        Ip and the remaining quotient by Carry. A non-empty address fragment
        is separated with a dot; an empty fragment receives no leading dot.

        Examples:
            IpBuildDoc(192, "168.1.1")
                = [Ip = "192.168.1.1", Carry = 0]
            IpBuildDoc(258, "1")
                = [Ip = "2.1", Carry = 1]

        Unlike VBA's implicit numeric coercion, the M contract requires a
        non-negative whole number that can be represented exactly by M's
        double-precision number type.
    */
    IpBuild = (ipByte as nullable number, ip as nullable text) as record =>
        let
            fail = IpMakeFailure("IpBuild.InvalidInput"),

            checkedIpByte =
                if ipByte = null then
                    fail(
                        "IPv4 byte value cannot be null.",
                        [Input = ipByte, Component = "byte", Expected = "non-negative whole number"]
                    )
                else if ipByte < 0 then
                    fail(
                        "IPv4 byte value must be non-negative.",
                        [Input = ipByte, Component = "byte", Expected = "non-negative whole number"]
                    )
                else if Number.RoundDown(ipByte) <> ipByte then
                    fail(
                        "IPv4 byte value must be a whole number.",
                        [Input = ipByte, Component = "byte", Expected = "non-negative whole number"]
                    )
                else if ipByte > Number.Power(2, 53) - 1 then
                    fail(
                        "IPv4 byte value must be within M's exact integer range.",
                        [Input = ipByte, Component = "byte", Expected = "0..9007199254740991"]
                    )
                else
                    ipByte,

            checkedIp =
                if ip = null then
                    fail(
                        "IPv4 address fragment cannot be null.",
                        [Input = ip, Component = "address fragment", Expected = "text"]
                    )
                else
                    ip,
            carry = Number.IntegerDivide(checkedIpByte, 256),
            lowByte = Number.Mod(checkedIpByte, 256),
            prefix = Number.ToText(lowByte, "0", "en-US"),
            builtIp = if checkedIp = "" then prefix else prefix & "." & checkedIp
        in
            [Ip = builtIp, Carry = carry],

    IpBuildType =
        type function (
           ipByte as (type nullable number meta [
                Documentation.FieldCaption = "IPv4 byte value",
                Documentation.FieldDescription = "A non-negative whole number whose low eight bits are prepended."
           ]),
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address fragment",
                Documentation.FieldDescription = "The existing dotted-decimal fragment to receive the new low byte."
           ])
        ) as [Ip = text, Carry = number] meta [
            Documentation.Name = "IpBuild",
            Documentation.LongDescription = "Returns a record with Ip and Carry fields. Ip contains the low eight bits followed by the original fragment, and Carry is the integer quotient after division by 256. This record is the immutable M equivalent of the VBA function's ByRef string and numeric return value.",
            Documentation.Examples = {
                [
                    Description = "Prepend a byte without a carry.",
                    Code = "IpBuild(192, \"168.1.1\")",
                    Result = "[Ip = \"192.168.1.1\", Carry = 0]"
                ],
                [
                    Description = "Prepend the low byte and return the carry.",
                    Code = "IpBuild(258, \"1\")",
                    Result = "[Ip = \"2.1\", Carry = 1]"
                ],
                [
                    Description = "Build the first byte of an address.",
                    Code = "IpBuild(255, \"\")",
                    Result = "[Ip = \"255\", Carry = 0]"
                ]
            }
        ],

    IpBuildDoc = Value.ReplaceType(IpBuild, IpBuildType),

    /*
        Adds a signed address offset to an IPv4 address.

        The VBA and JavaScript references both convert the address to its
        base-256 integer representation, add the offset, and format the
        result as dotted-decimal text. M uses the same composition through
        IpStrToBin and IpBinToStr. Offsets must be whole numbers; negative
        offsets are allowed when the resulting address remains in the
        unsigned IPv4 range.

        Examples:
            IpAddDoc("192.168.1.1", 4) = "192.168.1.5"
            IpAddDoc("192.168.1.1", 256) = "192.168.2.1"
            IpAddDoc("192.168.1.0", -1) = "192.168.0.255"

        Unlike the permissive reference conversions, null, fractional, and
        out-of-range results raise structured validation errors. The existing
        IpStrToBin and IpBinToStr helpers perform the address validation and
        final range check exactly once each.
    */
    IpAdd = (ip as nullable text, offset as nullable number) as text =>
        let
            fail = IpMakeFailure("IpAdd.InvalidInput")
        in
            if offset = null then
                fail(
                    "IPv4 offset cannot be null.",
                    [Input = offset, Component = "offset", Expected = "signed whole number"]
                )
            else if Number.Abs(offset) > Number.Power(2, 53) - 1 then
                fail(
                    "IPv4 offset must be within M's exact whole-number range.",
                    [Input = offset, Component = "offset", Expected = "signed whole number from -(2^53-1) through 2^53-1"]
                )
            else if Number.RoundDown(offset) <> offset then
                fail(
                    "IPv4 offset must be a whole number.",
                    [Input = offset, Component = "offset", Expected = "signed whole number"]
                )
            else
                let
                    addressValue = IpStrToBin(ip),
                    resultValue = addressValue + offset
                in
                    IpBinToStr(resultValue),

    IpAddType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           offset as (type nullable number meta [
                Documentation.FieldCaption = "Address offset",
                Documentation.FieldDescription = "A signed whole number of IPv4 addresses to add."
           ])
        ) as text meta [
            Documentation.Name = "IpAdd",
            Documentation.LongDescription = "Adds a signed whole-number offset within M's exact integer range to a validated IPv4 address and returns dotted-decimal text. The result must remain in the range 0 through 4294967295; null, fractional, inexact, malformed, and out-of-range inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Add four addresses.",
                    Code = "IpAdd(\"192.168.1.1\", 4)",
                    Result = "192.168.1.5"
                ],
                [
                    Description = "Carry across the final octet.",
                    Code = "IpAdd(\"192.168.1.1\", 256)",
                    Result = "192.168.2.1"
                ],
                [
                    Description = "Subtract one address.",
                    Code = "IpAdd(\"192.168.1.0\", -1)",
                    Result = "192.168.0.255"
                ]
            }
        ],

    IpAddDoc = Value.ReplaceType(IpAdd, IpAddType),

    /*
        Preserves the alternate VBA IpAdd2 compatibility name.

        IpAdd2 implements the same address operation as IpAdd with a
        right-to-left carry loop instead of converting the full address to a
        number first. For validated IPv4 text and a signed whole-number
        offset, both algorithms produce the same result. The M-native
        implementation therefore delegates to IpAdd so the public contract,
        range checks, and boundary behavior have one source of truth rather
        than duplicating arithmetic and validation.

        Examples:
            IpAdd2Doc("192.168.1.1", 4) = "192.168.1.5"
            IpAdd2Doc("192.168.1.1", 256) = "192.168.2.1"
            IpAdd2Doc("192.168.1.0", -1) = "192.168.0.255"
    */
    IpAdd2 = (ip as nullable text, offset as nullable number) as text =>
        IpAdd(ip, offset),

    IpAdd2Type =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           offset as (type nullable number meta [
                Documentation.FieldCaption = "Address offset",
                Documentation.FieldDescription = "A signed whole number of IPv4 addresses to add."
           ])
        ) as text meta [
            Documentation.Name = "IpAdd2",
            Documentation.LongDescription = "Preserves the VBA IpAdd2 function name while using the shared IpAdd contract. Adds a signed whole-number offset to a validated IPv4 address; the result must remain in the range 0 through 4294967295. Null, fractional, malformed, and out-of-range inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Add four addresses.",
                    Code = "IpAdd2(\"192.168.1.1\", 4)",
                    Result = "192.168.1.5"
                ],
                [
                    Description = "Carry across an octet boundary.",
                    Code = "IpAdd2(\"192.168.1.1\", 256)",
                    Result = "192.168.2.1"
                ],
                [
                    Description = "Subtract one address.",
                    Code = "IpAdd2(\"192.168.1.0\", -1)",
                    Result = "192.168.0.255"
                ]
            }
        ],

    IpAdd2Doc = Value.ReplaceType(IpAdd2, IpAdd2Type),

    /*
        Compares the first n bits of two IPv4 addresses.

        The VBA reference compares complete octets from left to right and,
        when n ends inside an octet, compares only that octet's leading bits.
        This M implementation uses the validated IPv4 integer representation
        and integer-dividing away the host bits. A zero-bit comparison therefore
        returns true for any two valid IPv4 addresses.

        The JavaScript reference has no direct public IpComp function. Its
        IpNet.matchIp and IpNet.matchSubnet operations provide the related
        subnet-prefix behavior, while this function preserves the VBA API's
        arbitrary prefix length.

        Examples:
            IpCompDoc("10.0.0.0", "10.1.0.0", 9) = true
            IpCompDoc("10.0.0.0", "10.1.0.0", 16) = false
            IpCompDoc("192.168.1.1", "10.0.0.1", 0) = true

        Unlike VBA's implicit Val conversion, both addresses must be valid
        dotted-decimal IPv4 text and n must be a whole number from 0 through
        32. Invalid inputs raise the structured validation errors from
        IpStrToBin or the explicit prefix validation below.
    */
    IpComp = (ip1 as nullable text, ip2 as nullable text, n as nullable number) as logical =>
        let
            fail = IpMakeFailure("IpComp.InvalidInput"),

            prefixLength =
                if n = null then
                    fail(
                        "Prefix length cannot be null.",
                        [Input = n, Component = "prefix length", Expected = "whole number 0..32"]
                    )
                else if n < 0 or n > 32 or Number.RoundDown(n) <> n then
                    fail(
                        "Prefix length must be a whole number from 0 through 32.",
                        [Input = n, Component = "prefix length", Expected = "whole number 0..32"]
                    )
                else
                    n,

            firstAddress = IpStrToBin(ip1),
            secondAddress = IpStrToBin(ip2),
            hostBitCount = 32 - prefixLength,
            divisor = Number.Power(2, hostBitCount),
            firstPrefix = Number.IntegerDivide(firstAddress, divisor),
            secondPrefix = Number.IntegerDivide(secondAddress, divisor)
        in
            firstPrefix = secondPrefix,

    IpCompType =
        type function (
           ip1 as (type nullable text meta [
                Documentation.FieldCaption = "First IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           ip2 as (type nullable text meta [
                Documentation.FieldCaption = "Second IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           n as (type nullable number meta [
                Documentation.FieldCaption = "Prefix length",
                Documentation.FieldDescription = "The number of leading bits to compare, from 0 through 32."
           ])
        ) as logical meta [
            Documentation.Name = "IpComp",
            Documentation.LongDescription = "Returns true when the first n bits of both validated dotted-decimal IPv4 addresses are equal. Prefix lengths from 0 through 32 are supported; a zero-bit comparison returns true. Invalid addresses or prefix lengths raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Compare a prefix that ends inside the second octet.",
                    Code = "IpComp(\"10.0.0.0\", \"10.1.0.0\", 9)",
                    Result = "true"
                ],
                [
                    Description = "Show that the same addresses differ in their first 16 bits.",
                    Code = "IpComp(\"10.0.0.0\", \"10.1.0.0\", 16)",
                    Result = "false"
                ],
                [
                    Description = "Compare zero leading bits.",
                    Code = "IpComp(\"192.168.1.1\", \"10.0.0.1\", 0)",
                    Result = "true"
                ]
            }
        ],

    IpCompDoc = Value.ReplaceType(IpComp, IpCompType),

    /*
        Performs a bitwise AND on two IPv4 addresses.

        The VBA reference parses both operands from right to left and applies
        bitwise AND to each byte. The JavaScript counterpart applies the same
        operation to the four split octets. M can express the equivalent
        calculation directly on the validated base-256 integer values with
        Number.BitwiseAnd, then format the result as dotted-decimal text.

        Both operands are IPv4 addresses, not a specially typed address and
        mask pair. A canonical-mask check would reject valid generic bitwise
        AND operations and is therefore left to IpMaskLen or subnet helpers
        that explicitly interpret an operand as a mask.

        Examples:
            IpAndDoc("192.168.1.1", "255.255.255.0") = "192.168.1.0"
            IpAndDoc("255.255.255.255", "128.0.0.1") = "128.0.0.1"
            IpAndDoc("0.0.0.0", "255.255.255.255") = "0.0.0.0"

        Null, malformed, and out-of-range inputs raise the structured
        validation errors from IpStrToBin rather than being coerced.
    */
    IpAnd = (ip1 as nullable text, ip2 as nullable text) as text =>
        let
            firstAddress = IpStrToBin(ip1),
            secondAddress = IpStrToBin(ip2),
            result = Number.BitwiseAnd(firstAddress, secondAddress)
        in
            IpBinToStr(result),

    IpAndType =
        type function (
           ip1 as (type nullable text meta [
                Documentation.FieldCaption = "First IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           ip2 as (type nullable text meta [
                Documentation.FieldCaption = "Second IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ])
        ) as text meta [
            Documentation.Name = "IpAnd",
            Documentation.LongDescription = "Returns the dotted-decimal IPv4 result of applying bitwise AND to two validated IPv4 addresses. Both operands are generic IPv4 values; the second operand is not required to be a canonical subnet mask. Invalid or null inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Apply a /24 mask to an IPv4 address.",
                    Code = "IpAnd(\"192.168.1.1\", \"255.255.255.0\")",
                    Result = "192.168.1.0"
                ],
                [
                    Description = "Preserve the high and low bits selected by a generic IPv4 operand.",
                    Code = "IpAnd(\"255.255.255.255\", \"128.0.0.1\")",
                    Result = "128.0.0.1"
                ],
                [
                    Description = "AND with the all-zero address.",
                    Code = "IpAnd(\"0.0.0.0\", \"255.255.255.255\")",
                    Result = "0.0.0.0"
                ]
            }
        ],

    IpAndDoc = Value.ReplaceType(IpAnd, IpAndType),

    /*
        Performs a bitwise OR on two IPv4 addresses.

        The VBA reference applies OR to corresponding octets from right to
        left, while the JavaScript counterpart applies OR to the four split
        octets. M can express the same operation directly on validated
        base-256 integer values with Number.BitwiseOr and then format the
        result as dotted-decimal text.

        Both operands are generic IPv4 addresses, not a specially typed
        address and mask pair. A canonical-mask check would change the
        reference operation's meaning and is therefore not performed here.

        Examples:
            IpOrDoc("192.168.1.1", "0.0.0.255") = "192.168.1.255"
            IpOrDoc("10.0.0.1", "0.0.255.0") = "10.0.255.1"
            IpOrDoc("0.0.0.0", "255.255.255.255") = "255.255.255.255"

        Null, malformed, and out-of-range inputs raise the structured
        validation errors from IpStrToBin rather than being coerced.
    */
    IpOr = (ip1 as nullable text, ip2 as nullable text) as text =>
        let
            firstAddress = IpStrToBin(ip1),
            secondAddress = IpStrToBin(ip2),
            result = Number.BitwiseOr(firstAddress, secondAddress)
        in
            IpBinToStr(result),

    IpOrType =
        type function (
           ip1 as (type nullable text meta [
                Documentation.FieldCaption = "First IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           ip2 as (type nullable text meta [
                Documentation.FieldCaption = "Second IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ])
        ) as text meta [
            Documentation.Name = "IpOr",
            Documentation.LongDescription = "Returns the dotted-decimal IPv4 result of applying bitwise OR to two validated IPv4 addresses. Both operands are generic IPv4 values; the second operand is not required to be a canonical subnet mask. Invalid or null inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Set the low byte of an IPv4 address.",
                    Code = "IpOr(\"192.168.1.1\", \"0.0.0.255\")",
                    Result = "192.168.1.255"
                ],
                [
                    Description = "Set the third-octet bits selected by a generic IPv4 operand.",
                    Code = "IpOr(\"10.0.0.1\", \"0.0.255.0\")",
                    Result = "10.0.255.1"
                ],
                [
                    Description = "OR with the all-ones address.",
                    Code = "IpOr(\"0.0.0.0\", \"255.255.255.255\")",
                    Result = "255.255.255.255"
                ]
            }
        ],

    IpOrDoc = Value.ReplaceType(IpOr, IpOrType),

    /*
        Performs a bitwise XOR on two IPv4 addresses.

        The VBA reference applies XOR to corresponding octets from right to
        left, while the JavaScript counterpart applies XOR to the four split
        octets. M can express the same operation directly on validated
        base-256 integer values with Number.BitwiseXor and then format the
        result as dotted-decimal text.

        Both operands are generic IPv4 addresses, not a specially typed
        address and mask pair. A canonical-mask check would change the
        reference operation's meaning and is therefore not performed here.

        Examples:
            IpXorDoc("192.168.1.1", "0.0.0.255") = "192.168.1.254"
            IpXorDoc("255.255.255.255", "128.0.0.1") = "127.255.255.254"
            IpXorDoc("0.0.0.0", "255.255.255.255") = "255.255.255.255"

        Null, malformed, and out-of-range inputs raise the structured
        validation errors from IpStrToBin rather than being coerced.
    */
    IpXor = (ip1 as nullable text, ip2 as nullable text) as text =>
        let
            firstAddress = IpStrToBin(ip1),
            secondAddress = IpStrToBin(ip2),
            result = Number.BitwiseXor(firstAddress, secondAddress)
        in
            IpBinToStr(result),

    IpXorType =
        type function (
           ip1 as (type nullable text meta [
                Documentation.FieldCaption = "First IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           ip2 as (type nullable text meta [
                Documentation.FieldCaption = "Second IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ])
        ) as text meta [
            Documentation.Name = "IpXor",
            Documentation.LongDescription = "Returns the dotted-decimal IPv4 result of applying bitwise XOR to two validated IPv4 addresses. Both operands are generic IPv4 values; the second operand is not required to be a canonical subnet mask. Invalid or null inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Toggle the low byte of an IPv4 address.",
                    Code = "IpXor(\"192.168.1.1\", \"0.0.0.255\")",
                    Result = "192.168.1.254"
                ],
                [
                    Description = "XOR a generic all-ones operand with selected bits.",
                    Code = "IpXor(\"255.255.255.255\", \"128.0.0.1\")",
                    Result = "127.255.255.254"
                ],
                [
                    Description = "XOR with the all-zero address.",
                    Code = "IpXor(\"0.0.0.0\", \"255.255.255.255\")",
                    Result = "255.255.255.255"
                ]
            }
        ],

    IpXorDoc = Value.ReplaceType(IpXor, IpXorType),

    /*
        Returns the signed difference between two IPv4 addresses.

        The VBA reference subtracts the base-256 bytes from right to left,
        and the JavaScript counterpart subtracts the two numeric IPv4
        representations. M can express the same operation directly after
        parsing each validated address once with IpStrToBin.

        Examples:
            IpDiffDoc("192.168.1.7", "192.168.1.1") = 6
            IpDiffDoc("192.168.2.1", "192.168.1.1") = 256
            IpDiffDoc("192.168.1.1", "192.168.1.7") = -6

        Null or malformed inputs raise the structured validation errors from
        IpStrToBin instead of being coerced as in the permissive references.
    */
    IpDiff = (ip1 as nullable text, ip2 as nullable text) as number =>
        let
            firstAddress = IpStrToBin(ip1),
            secondAddress = IpStrToBin(ip2)
        in
            firstAddress - secondAddress,

    IpDiffType =
        type function (
           ip1 as (type nullable text meta [
                Documentation.FieldCaption = "First IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           ip2 as (type nullable text meta [
                Documentation.FieldCaption = "Second IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ])
        ) as number meta [
            Documentation.Name = "IpDiff",
            Documentation.LongDescription = "Subtracts the numeric representation of the second validated IPv4 address from the first and returns the signed difference in address units. Invalid or null inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Find the difference within one final octet.",
                    Code = "IpDiff(\"192.168.1.7\", \"192.168.1.1\")",
                    Result = "6"
                ],
                [
                    Description = "Carry across an octet boundary.",
                    Code = "IpDiff(\"192.168.2.1\", \"192.168.1.1\")",
                    Result = "256"
                ],
                [
                    Description = "Return a negative difference when the first address is lower.",
                    Code = "IpDiff(\"192.168.1.1\", \"192.168.1.7\")",
                    Result = "-6"
                ]
            }
        ],

    IpDiffDoc = Value.ReplaceType(IpDiff, IpDiffType),

    /* Converts a validated IPv4 integer to four left-to-right octets. */
    IpAddressValueToOctets = (addressValue as number) as list =>
        List.Transform(
            {3, 2, 1, 0},
            (power as number) as number =>
                Number.Mod(Number.IntegerDivide(addressValue, Number.Power(256, power)), 256)
        ),

    /*
        Replaces one 1-based IPv4 octet from left to right.

        The VBA and JavaScript references both use a one-based position: 1 is
        the first octet and 4 is the final octet. M first validates the
        address through IpStrToBin, derives the four octets from that parsed
        integer, replaces exactly one list item, and formats the result back
        to canonical dotted-decimal text.

        The position and replacement value must be whole numbers in the
        ranges 1..4 and 0..255 respectively. Null, malformed, fractional,
        and out-of-range inputs raise a structured validation error rather
        than inheriting the permissive reference coercion behavior.

        Examples:
            IpSetByteDoc("192.168.1.1", 4, 20) = "192.168.1.20"
            IpSetByteDoc("192.168.1.1", 1, 10) = "10.168.1.1"
            IpSetByteDoc("0.0.0.0", 2, 255) = "0.255.0.0"
    */
    IpSetByte = (
        ip as nullable text,
        position as nullable number,
        newValue as nullable number
    ) as text =>
        let
            fail = IpMakeFailure("IpSetByte.InvalidInput")
        in
            if position = null then
                fail(
                    "IPv4 octet position cannot be null.",
                    [Input = position, Component = "position", Expected = "whole number 1..4"]
                )
            else if position < 1 or position > 4 or Number.RoundDown(position) <> position then
                fail(
                    "IPv4 octet position must be a whole number from 1 through 4.",
                    [Input = position, Component = "position", Expected = "whole number 1..4"]
                )
            else if newValue = null then
                fail(
                    "Replacement IPv4 octet cannot be null.",
                    [Input = newValue, Component = "newValue", Expected = "whole number 0..255"]
                )
            else if newValue < 0 or newValue > 255 or Number.RoundDown(newValue) <> newValue then
                fail(
                    "Replacement IPv4 octet must be a whole number from 0 through 255.",
                    [Input = newValue, Component = "newValue", Expected = "whole number 0..255"]
                )
            else
                let
                    addressValue = IpStrToBin(ip),
                    octets = IpAddressValueToOctets(addressValue),
                    replacedOctets = List.ReplaceRange(octets, position - 1, 1, {newValue})
                in
                    Text.Combine(
                        List.Transform(
                            replacedOctets,
                            (octet as number) as text => Number.ToText(octet, "0", "en-US")
                        ),
                        "."
                    ),

    IpSetByteType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           position as (type nullable number meta [
                Documentation.FieldCaption = "Octet position",
                Documentation.FieldDescription = "A one-based octet position from 1 through 4, counted from left to right."
           ]),
           newValue as (type nullable number meta [
                Documentation.FieldCaption = "Replacement octet",
                Documentation.FieldDescription = "A whole-number IPv4 octet from 0 through 255."
           ])
        ) as text meta [
            Documentation.Name = "IpSetByte",
            Documentation.LongDescription = "Replaces the one-based octet at position 1 through 4 in a validated IPv4 address and returns canonical dotted-decimal text. Null, malformed, fractional, and out-of-range inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Replace the final octet.",
                    Code = "IpSetByte(\"192.168.1.1\", 4, 20)",
                    Result = "192.168.1.20"
                ],
                [
                    Description = "Replace the first octet.",
                    Code = "IpSetByte(\"192.168.1.1\", 1, 10)",
                    Result = "10.168.1.1"
                ],
                [
                    Description = "Replace an interior octet with the maximum value.",
                    Code = "IpSetByte(\"0.0.0.0\", 2, 255)",
                    Result = "0.255.0.0"
                ]
            }
        ],

    IpSetByteDoc = Value.ReplaceType(IpSetByte, IpSetByteType),

    /*
        Returns one 1-based IPv4 octet from left to right.

        The VBA and JavaScript references both use positions 1 through 4,
        where position 1 is the first octet. M validates the address once
        through IpStrToBin, derives the four octets with exact base-256
        arithmetic, and selects the requested list item. This avoids the
        VBA helper's mutable right-to-left parsing state while preserving its
        public result for valid canonical or leading-zero IPv4 text.

        The position must be a whole number from 1 through 4. Null,
        malformed, fractional, and out-of-range inputs raise structured
        validation errors rather than inheriting reference coercion behavior.

        Examples:
            IpGetByteDoc("192.168.1.1", 1) = 192
            IpGetByteDoc("192.168.1.1", 3) = 1
            IpGetByteDoc("255.0.0.1", 4) = 1
    */
    IpGetByte = (
        ip as nullable text,
        position as nullable number
    ) as number =>
        let
            fail = IpMakeFailure("IpGetByte.InvalidInput")
        in
            if position = null then
                fail(
                    "IPv4 octet position cannot be null.",
                    [Input = position, Component = "position", Expected = "whole number 1..4"]
                )
            else if position < 1 or position > 4 or Number.RoundDown(position) <> position then
                fail(
                    "IPv4 octet position must be a whole number from 1 through 4.",
                    [Input = position, Component = "position", Expected = "whole number 1..4"]
                )
            else
                IpAddressValueToOctets(IpStrToBin(ip)){position - 1},

    IpGetByteType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           position as (type nullable number meta [
                Documentation.FieldCaption = "Octet position",
                Documentation.FieldDescription = "A one-based octet position from 1 through 4, counted from left to right."
           ])
        ) as number meta [
            Documentation.Name = "IpGetByte",
            Documentation.LongDescription = "Returns the one-based octet at position 1 through 4 in a validated IPv4 address. Null, malformed, fractional, and out-of-range inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Get the first octet.",
                    Code = "IpGetByte(\"192.168.1.1\", 1)",
                    Result = "192"
                ],
                [
                    Description = "Get an interior octet.",
                    Code = "IpGetByte(\"192.168.1.1\", 3)",
                    Result = "1"
                ],
                [
                    Description = "Get the final octet.",
                    Code = "IpGetByte(\"255.0.0.1\", 4)",
                    Result = "1"
                ]
            }
        ],

    IpGetByteDoc = Value.ReplaceType(IpGetByte, IpGetByteType),

    /*
        Sorts a list of IPv4 addresses numerically.

        The VBA function accepts a single-column Range, converts each address
        to its base-256 integer value, sorts those values, and exports the
        result as a single-column array. The JavaScript counterpart performs
        the same numeric sort in ascending order. M represents the imported
        range as a list, ignores null and empty text cells, and returns a
        canonical dotted-decimal address list. The optional descending flag
        preserves VBA's additional ordering behavior.

        Each non-empty address is parsed once before sorting. Remaining items
        must be text and must satisfy the strict IPv4 address contract.

        Examples:
            IpSortArrayDoc({"192.168.1.2", "10.0.0.1"})
                = {"10.0.0.1", "192.168.1.2"}
            IpSortArrayDoc({"192.168.1.2", "10.0.0.1"}, true)
                = {"192.168.1.2", "10.0.0.1"}
            IpSortArrayDoc({"010.000.000.001", "10.0.0.2"})
                = {"10.0.0.1", "10.0.0.2"}
    */
    IpSortArray = (
        addresses as nullable list,
        optional descending as nullable logical
    ) as list =>
        let
            fail = IpMakeFailure("IpSortArray.InvalidInput"),

            checkedAddresses =
                if addresses = null then
                    fail(
                        "IPv4 address list cannot be null.",
                        [Input = addresses, Component = "addresses", Expected = "list of IPv4 address text"]
                    )
                else
                    addresses,
            nonEmptyAddresses =
                List.Select(
                    checkedAddresses,
                    (rawAddress as any) as logical =>
                        if rawAddress = null then
                            false
                        else if Value.Is(rawAddress, type text) then
                            rawAddress <> ""
                        else
                            true
                ),
            parsedAddresses =
                List.Buffer(
                    List.Transform(
                        nonEmptyAddresses,
                        (rawAddress as any) as number =>
                            if not Value.Is(rawAddress, type text) then
                                fail(
                                    "IPv4 address list items must be text.",
                                    [Input = rawAddress, Component = "address", Expected = "text"]
                                )
                            else
                                IpStrToBin(rawAddress)
                    )
                ),
            descendingMode = descending ?? false,
            sortedAddresses =
                List.Sort(
                    parsedAddresses,
                    if descendingMode then Order.Descending else Order.Ascending
                )
        in
            List.Transform(sortedAddresses, (addressValue as number) as text => IpBinToStr(addressValue)),

    IpSortArrayType =
        type function (
           addresses as (type nullable list meta [
                Documentation.FieldCaption = "IPv4 address list",
                Documentation.FieldDescription = "A list of dotted IPv4 address text values; null and empty text items are ignored."
           ]),
           optional descending as (type nullable logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Return the sorted address list in descending numeric order."
           ])
        ) as list meta [
            Documentation.Name = "IpSortArray",
            Documentation.LongDescription = "Sorts a Power Query list of IPv4 address text by numeric address value and returns canonical dotted-decimal text. Null and empty text items are ignored, non-text or malformed address items raise structured validation errors, and a null descending value means ascending order.",
            Documentation.Examples = {
                [
                    Description = "Sort addresses in ascending numeric order.",
                    Code = "IpSortArray({\"192.168.1.2\", \"10.0.0.1\"})",
                    Result = "{\"10.0.0.1\", \"192.168.1.2\"}"
                ],
                [
                    Description = "Sort addresses in descending numeric order.",
                    Code = "IpSortArray({\"192.168.1.2\", \"10.0.0.1\"}, true)",
                    Result = "{\"192.168.1.2\", \"10.0.0.1\"}"
                ],
                [
                    Description = "Canonicalize leading-zero address text while sorting.",
                    Code = "IpSortArray({\"010.000.000.001\", \"10.0.0.2\"})",
                    Result = "{\"10.0.0.1\", \"10.0.0.2\"}"
                ]
            }
        ],

    IpSortArrayDoc = Value.ReplaceType(IpSortArray, IpSortArrayType),

    /*
        Inverts every bit of a dotted-decimal IPv4 mask or wildcard mask.

        The VBA and JavaScript references subtract the parsed IPv4 value from
        the all-ones 32-bit value. This maps a subnet mask to its wildcard
        mask and maps a wildcard mask back to its subnet mask. M preserves the
        generic complement operation rather than requiring a canonical
        contiguous netmask; non-contiguous IPv4 bit patterns are valid inputs
        to this operation even though IpMaskLen deliberately rejects them.

        Examples:
            IpInvertMaskDoc("255.255.255.0") = "0.0.0.255"
            IpInvertMaskDoc("0.0.0.255") = "255.255.255.0"
            IpInvertMaskDoc("255.0.255.0") = "0.255.0.255"

        Null or malformed inputs raise the structured validation errors from
        IpStrToBin instead of being coerced by the permissive references.
    */
    IpInvertMask = (mask as nullable text) as text =>
        let
            maskValue = IpStrToBin(mask),
            allBits = Number.Power(2, 32) - 1,
            invertedValue = allBits - maskValue
        in
            IpBinToStr(invertedValue),

    IpInvertMaskType =
        type function (
           mask as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 mask or wildcard mask",
                Documentation.FieldDescription = "A dotted-decimal IPv4 value whose 32 bits should be inverted."
           ])
        ) as text meta [
            Documentation.Name = "IpInvertMask",
            Documentation.LongDescription = "Returns the 32-bit complement of a validated dotted-decimal IPv4 value. This converts canonical subnet masks to wildcard masks and wildcard masks to subnet masks, while also accepting non-contiguous bit patterns as generic IPv4 values. Null or malformed inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Invert a /24 subnet mask.",
                    Code = "IpInvertMask(\"255.255.255.0\")",
                    Result = "0.0.0.255"
                ],
                [
                    Description = "Invert a /24 wildcard mask.",
                    Code = "IpInvertMask(\"0.0.0.255\")",
                    Result = "255.255.255.0"
                ],
                [
                    Description = "Invert a non-contiguous IPv4 bit pattern.",
                    Code = "IpInvertMask(\"255.0.255.0\")",
                    Result = "0.255.0.255"
                ]
            }
        ],

    IpInvertMaskDoc = Value.ReplaceType(IpInvertMask, IpInvertMaskType),

    /* Returns normalized numeric facts shared by subnet-list operations. */
    IpSubnetFacts = (addressValue as number, prefixLength as number) as [NetworkKey = number, PrefixLength = number, BlockSize = number] =>
        let
            blockSize = Number.Power(2, 32 - prefixLength),
            networkKey = Number.IntegerDivide(addressValue, blockSize) * blockSize
        in
            [NetworkKey = networkKey, PrefixLength = prefixLength, BlockSize = blockSize],

    /* Renders one normalized candidate without reparsing subnet text. */
    IpSubnetCandidateFromFacts = (
        networkKey as number,
        prefixLength as number,
        usesCidr as logical,
        sourcePosition as number
    ) as record =>
        let
            facts = IpSubnetFacts(networkKey, prefixLength),
            networkText = IpBinToStr(facts[NetworkKey]),
            maskValue = (Number.Power(2, prefixLength) - 1) * facts[BlockSize],
            output = if usesCidr then networkText & "/" & Number.ToText(prefixLength, "0", "en-US") else networkText & " " & IpBinToStr(maskValue)
        in
            [NetworkKey = facts[NetworkKey], PrefixLength = prefixLength, UsesCidr = usesCidr, SourcePosition = sourcePosition, Output = output],

    /* Parses and canonicalizes one candidate for all subnet-list operations. */
    IpSubnetCandidateFromValue = (rawSubnet as any, sourcePosition as number, errorName as text) as record =>
        let
            checkedSubnet =
                if not Value.Is(rawSubnet, type text) then
                    error Error.Record(errorName, "Subnet list items must contain text.", [Input = rawSubnet, Component = "subnet", Expected = "IPv4 subnet text"])
                else rawSubnet,
            parsedSubnet = IpSubnetParseWithAddressValue(checkedSubnet),
            facts = parsedSubnet[Facts],
            usesCidr = Text.PositionOf(checkedSubnet, "/") >= 0 or Text.PositionOf(checkedSubnet, " ") < 0
        in
            IpSubnetCandidateFromFacts(facts[NetworkKey], facts[PrefixLength], usesCidr, sourcePosition),

    /* Removes null/empty cells while preserving invalid values for validation. */
    IpSubnetCleanList = (values as nullable list, listName as text, errorName as text) as list =>
        let
            checkedValues =
                if values = null then
                    error Error.Record(errorName, listName & " list cannot be null.", [Input = values, Component = listName, Expected = "list of subnet text"])
                else values
        in
            List.Select(checkedValues, (rawSubnet as any) as logical => if rawSubnet = null then false else if Value.Is(rawSubnet, type text) then rawSubnet <> "" else true),

    /* Parses candidates in source order so notation tie-breaks remain stable. */
    IpSubnetCandidatesFromList = (values as list, errorName as text) as list =>
        List.Transform(List.Positions(values), (position as number) => IpSubnetCandidateFromValue(values{position}, position, errorName)),

    /* Tests normalized subnet containment without reparsing candidate text. */
    IpSubnetCandidateContains = (outer as record, inner as record) as logical =>
        inner[PrefixLength] >= outer[PrefixLength]
            and IpSubnetFacts(inner[NetworkKey], outer[PrefixLength])[NetworkKey] = outer[NetworkKey],

    /* Sorts by network and prefix, then source order for stable notation. */
    IpSubnetCandidateCompare = (left as record, right as record) as number =>
        if left[NetworkKey] < right[NetworkKey] then -1
        else if left[NetworkKey] > right[NetworkKey] then 1
        else if left[PrefixLength] < right[PrefixLength] then -1
        else if left[PrefixLength] > right[PrefixLength] then 1
        else Value.Compare(left[SourcePosition], right[SourcePosition]),

    /* Joins siblings and preserves notation from the earliest input candidate. */
    IpSubnetJoinedCandidate = (left as record, right as record) as nullable record =>
        if left[PrefixLength] <> right[PrefixLength] or left[PrefixLength] = 0 then null
        else
            let
                parentFacts = IpSubnetFacts(left[NetworkKey], left[PrefixLength] - 1),
                rightParentFacts = IpSubnetFacts(right[NetworkKey], right[PrefixLength] - 1),
                earlier = if left[SourcePosition] <= right[SourcePosition] then left else right
            in
                if parentFacts[NetworkKey] = rightParentFacts[NetworkKey] then
                    IpSubnetCandidateFromFacts(parentFacts[NetworkKey], parentFacts[PrefixLength], earlier[UsesCidr], earlier[SourcePosition])
                else null,

    /* A stack fold bounds merge recursion by the 32 IPv4 prefix levels. */
    IpSubnetAggregateCandidates = (candidates as list) as list =>
        let
            push = (stack as list, candidate as record) as list =>
                if List.IsEmpty(stack) then {candidate}
                else
                    let
                        previous = List.Last(stack),
                        joined = IpSubnetJoinedCandidate(previous, candidate)
                    in
                        if IpSubnetCandidateContains(previous, candidate) then stack
                        else if joined <> null then @push(List.RemoveLastN(stack, 1), joined)
                        else stack & {candidate}
        in
            List.Accumulate(List.Sort(candidates, IpSubnetCandidateCompare), {}, (stack as list, candidate as record) => push(stack, candidate)),

    /*
        Sorts and summarizes a list of IPv4 subnets.

        The VBA and JavaScript references remove empty cells, normalize each
        subnet to its network boundary, sort by network address and prefix,
        remove subnets contained by an earlier entry, and repeatedly merge
        aligned adjacent subnets of the same size. M represents the working
        values as records so each normalized address, prefix, and rendered
        result can be reused without reproducing Excel range mutation.

        The Power Query contract is a list of subnet text values. Null and
        empty text items are ignored as the reference import helpers ignore
        empty cells; every remaining item must be text. The result is a list,
        not a spilled worksheet array. Output addresses are canonical network
        addresses. CIDR input and unmasked input use CIDR output, while dotted
        mask input uses a canonical dotted mask. A null descending value means
        ascending order. When normalized records are deduplicated or merged,
        the earliest non-empty input involved determines the output notation.

        Examples:
            IpSubnetAggregateArrayDoc({"192.168.1.0/25", "192.168.1.128/25"})
                = {"192.168.1.0/24"}
            IpSubnetAggregateArrayDoc({"10.0.0.0/8", "10.1.0.0/16"})
                = {"10.0.0.0/8"}
            IpSubnetAggregateArrayDoc(
                {"192.168.1.0 255.255.255.128", "192.168.1.128 255.255.255.128"}
            ) = {"192.168.1.0 255.255.255.0"}
    */
    IpSubnetAggregateArray = (subnets as nullable list, optional descending as nullable logical) as list =>
        let
            nonEmptySubnets = IpSubnetCleanList(subnets, "subnets", "IpSubnetAggregateArray.InvalidInput"),
            candidates = IpSubnetCandidatesFromList(nonEmptySubnets, "IpSubnetAggregateArray.InvalidInput"),
            ascendingResult = List.Transform(IpSubnetAggregateCandidates(candidates), (candidate as record) => candidate[Output])
        in
            if descending ?? false then List.Reverse(ascendingResult) else ascendingResult,

    IpSubnetAggregateArrayType =
        type function (
           subnets as (type nullable list meta [
                Documentation.FieldCaption = "IPv4 subnet list",
                Documentation.FieldDescription = "A list of dotted IPv4 subnet text values; null and empty text items are ignored."
           ]),
           optional descending as (type nullable logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Return the aggregated list in descending network order."
           ])
        ) as list meta [
            Documentation.Name = "IpSubnetAggregateArray",
            Documentation.LongDescription = "Normalizes, sorts, deduplicates, removes contained subnets, and joins aligned adjacent equal-size IPv4 subnets. Returns a Power Query list of canonical network text; CIDR and unmasked inputs render as CIDR, dotted-mask inputs render with canonical dotted masks. Deduplicated and merged records retain notation from the earliest contributing non-empty input. Null or empty list items are ignored, non-text items and invalid subnet text raise structured validation errors, and a null descending value means ascending order.",
            Documentation.Examples = {
                [
                    Description = "Join two adjacent /25 networks.",
                    Code = "IpSubnetAggregateArray({\"192.168.1.0/25\", \"192.168.1.128/25\"})",
                    Result = "{\"192.168.1.0/24\"}"
                ],
                [
                    Description = "Remove a /16 subnet contained by a /8 network.",
                    Code = "IpSubnetAggregateArray({\"10.0.0.0/8\", \"10.1.0.0/16\"})",
                    Result = "{\"10.0.0.0/8\"}"
                ],
                [
                    Description = "Join dotted-mask input and retain dotted-mask output.",
                    Code = "IpSubnetAggregateArray({\"192.168.1.0 255.255.255.128\", \"192.168.1.128 255.255.255.128\"})",
                    Result = "{\"192.168.1.0 255.255.255.0\"}"
                ],
                [
                    Description = "Reverse the final aggregated list.",
                    Code = "IpSubnetAggregateArray({\"10.0.0.0/24\", \"10.0.2.0/24\"}, true)",
                    Result = "{\"10.0.2.0/24\", \"10.0.0.0/24\"}"
                ]
            }
        ],

    IpSubnetAggregateArrayDoc = Value.ReplaceType(IpSubnetAggregateArray, IpSubnetAggregateArrayType),

    /*
        Sorts a list of IPv4 subnets by network address and prefix length.

        The VBA helper sorts on a normalized network key and prefix while its
        range-export path returns the original route text. The JavaScript
        counterpart constructs IpNet values before sorting, so it returns
        canonical network text and clears host bits. M follows that public
        JavaScript behavior: the shortest prefix comes first when normalized
        network addresses are equal, unmasked input renders as /32, and
        dotted-mask input renders with a canonical dotted mask.

        The Power Query contract is a list of subnet text values. Null and
        empty text items are ignored as empty range cells are ignored by the
        reference import helpers; every remaining item must be text. A null
        descending value means ascending order. Equal normalized network and
        prefix keys retain the first non-empty input's notation.

        Examples:
            IpSubnetSortArrayDoc({"192.168.1.193/24", "10.0.0.1/8"})
                = {"10.0.0.0/8", "192.168.1.0/24"}
            IpSubnetSortArrayDoc(
                {"192.168.1.0/24", "10.0.0.0/8"},
                true
            ) = {"192.168.1.0/24", "10.0.0.0/8"}
    */
    IpSubnetSortArray = (subnets as nullable list, optional descending as nullable logical) as list =>
        let
            nonEmptySubnets = IpSubnetCleanList(subnets, "subnets", "IpSubnetSortArray.InvalidInput"),
            candidates = IpSubnetCandidatesFromList(nonEmptySubnets, "IpSubnetSortArray.InvalidInput"),
            sortedResult = List.Transform(List.Sort(candidates, IpSubnetCandidateCompare), (candidate as record) => candidate[Output])
        in
            if descending ?? false then List.Reverse(sortedResult) else sortedResult,

    IpSubnetSortArrayType =
        type function (
           subnets as (type nullable list meta [
                Documentation.FieldCaption = "IPv4 subnet list",
                Documentation.FieldDescription = "A list of dotted IPv4 subnet text values; null and empty text items are ignored."
           ]),
           optional descending as (type nullable logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Return the sorted list in descending network order."
           ])
        ) as list meta [
            Documentation.Name = "IpSubnetSortArray",
            Documentation.LongDescription = "Normalizes and sorts IPv4 subnet text by ascending network address and then ascending prefix length, with shorter prefixes first for the same network. Returns canonical network text; unmasked input renders as /32 and dotted-mask input uses a canonical dotted mask. Null or empty list items are ignored, non-text items and invalid subnet text raise structured validation errors, and a null descending value means ascending order.",
            Documentation.Examples = {
                [
                    Description = "Sort by normalized network address.",
                    Code = "IpSubnetSortArray({\"192.168.1.193/24\", \"10.0.0.1/8\"})",
                    Result = "{\"10.0.0.0/8\", \"192.168.1.0/24\"}"
                ],
                [
                    Description = "Reverse the sorted result.",
                    Code = "IpSubnetSortArray({\"192.168.1.0/24\", \"10.0.0.0/8\"}, true)",
                    Result = "{\"192.168.1.0/24\", \"10.0.0.0/8\"}"
                ],
                [
                    Description = "Sort equal network addresses by prefix length.",
                    Code = "IpSubnetSortArray({\"10.0.0.0/16\", \"10.0.0.0/8\"})",
                    Result = "{\"10.0.0.0/8\", \"10.0.0.0/16\"}"
                ]
            }
        ],

    IpSubnetSortArrayDoc = Value.ReplaceType(IpSubnetSortArray, IpSubnetSortArrayType),

    /* Subtracts candidate records without reparsing or mutating one shared list. */
    IpSubtractSubnetCandidates = (inputCandidates as list, subtractCandidates as list, fastMode as logical) as list =>
        let
            preparedInput = if fastMode then IpSubnetAggregateCandidates(inputCandidates) else inputCandidates,
            preparedSubtract = if fastMode then IpSubnetAggregateCandidates(subtractCandidates) else subtractCandidates,
            subtractOne = (inputCandidate as record, subtractCandidate as record) as list =>
                if IpSubnetCandidateContains(subtractCandidate, inputCandidate) then {}
                else if not IpSubnetCandidateContains(inputCandidate, subtractCandidate) then {inputCandidate}
                else
                    let
                        siblingPrefixes = {(inputCandidate[PrefixLength] + 1)..subtractCandidate[PrefixLength]},
                        siblings = List.Transform(
                            siblingPrefixes,
                            (prefixLength as number) as record =>
                                let
                                    childFacts = IpSubnetFacts(subtractCandidate[NetworkKey], prefixLength),
                                    parentFacts = IpSubnetFacts(subtractCandidate[NetworkKey], prefixLength - 1),
                                    siblingNetworkKey = if childFacts[NetworkKey] = parentFacts[NetworkKey] then childFacts[NetworkKey] + childFacts[BlockSize] else parentFacts[NetworkKey]
                                in
                                    IpSubnetCandidateFromFacts(siblingNetworkKey, prefixLength, true, inputCandidate[SourcePosition])
                        )
                    in
                        List.Sort(siblings, IpSubnetCandidateCompare),
            subtractFromFragments = (fragments as list, subtractCandidate as record) as list =>
                List.Combine(List.Transform(fragments, (fragment as record) => subtractOne(fragment, subtractCandidate))),
            subtractFromInput = (inputCandidate as record) as list =>
                List.Accumulate(preparedSubtract, {inputCandidate}, subtractFromFragments)
        in
            List.Combine(List.Transform(preparedInput, subtractFromInput)),

    /*
        Removes subnets from an IPv4 subnet list.

        The VBA and JavaScript references remove each subtracting subnet from
        the input list. When a subtractor is contained by an input subnet, the
        result is the ordered set of sibling CIDRs along the path to that
        subtractor, equivalent to repeatedly splitting into next-prefix
        children. When the input subnet is contained by the subtractor, it is
        removed.

        The Power Query contract replaces Excel ranges and spilled arrays with
        two lists of subnet text. Null and empty text items are ignored as
        empty range cells; every remaining item must be text. Fast mode is the
        default and first aggregates both candidate lists without reparsing
        their rendered text. Fast=false preserves input order as far as removal
        and splitting allow, matching the reference's slow mode. A null fast
        value means true. Empty input or subtract lists are valid; an empty
        input returns empty, while an empty subtract list returns the normalized
        input in fast mode and canonicalized input order in slow mode.

        Examples:
            IpSubtractSubnetsDoc(
                {"10.0.0.0/24"},
                {"10.0.0.0/25"}
            ) = {"10.0.0.128/25"}
            IpSubtractSubnetsDoc(
                {"10.0.0.0/24"},
                {"10.0.1.0/24"}
            ) = {"10.0.0.0/24"}
            IpSubtractSubnetsDoc(
                {"10.0.0.0/24", "192.0.2.0/24"},
                {"10.0.0.128/25"},
                false
            ) = {"10.0.0.0/25", "192.0.2.0/24"}
    */
    IpSubtractSubnets = (inputSubnets as nullable list, subtractSubnets as nullable list, optional fast as nullable logical) as list =>
        let
            errorName = "IpSubtractSubnets.InvalidInput",
            checkedInputSubnets = IpSubnetCleanList(inputSubnets, "input", errorName),
            checkedSubtractSubnets = IpSubnetCleanList(subtractSubnets, "subtract", errorName),
            inputCandidates = IpSubnetCandidatesFromList(checkedInputSubnets, errorName),
            subtractCandidates = IpSubnetCandidatesFromList(checkedSubtractSubnets, errorName),
            resultCandidates = IpSubtractSubnetCandidates(inputCandidates, subtractCandidates, fast ?? true)
        in
            List.Transform(resultCandidates, (candidate as record) => candidate[Output]),

    IpSubtractSubnetsType =
        type function (
           inputSubnets as (type nullable list meta [
                Documentation.FieldCaption = "Input IPv4 subnets",
                Documentation.FieldDescription = "A list of assigned IPv4 subnet text values; null and empty text items are ignored."
           ]),
           subtractSubnets as (type nullable list meta [
                Documentation.FieldCaption = "Subtraction IPv4 subnets",
                Documentation.FieldDescription = "A list of IPv4 subnet text values to remove; null and empty text items are ignored."
           ]),
           optional fast as (type nullable logical meta [
                Documentation.FieldCaption = "Fast mode",
                Documentation.FieldDescription = "Aggregate and sort both lists before subtracting; null means true."
           ])
        ) as list meta [
            Documentation.Name = "IpSubtractSubnets",
            Documentation.LongDescription = "Removes each subtracting subnet from an input IPv4 subnet list. Fast mode, enabled by default, aggregates both lists and uses sorted network keys; slow mode preserves input order as far as removal and splitting allow. A containing input subnet is replaced by the ordered sibling CIDRs outside the subtractor. Returns a Power Query list, ignores null or empty cells, and raises structured validation errors for non-text or malformed subnet values.",
            Documentation.Examples = {
                [
                    Description = "Subtract the lower half of a /24.",
                    Code = "IpSubtractSubnets({\"10.0.0.0/24\"}, {\"10.0.0.0/25\"})",
                    Result = "{\"10.0.0.128/25\"}"
                ],
                [
                    Description = "Leave an input subnet unchanged when the subtractor is disjoint.",
                    Code = "IpSubtractSubnets({\"10.0.0.0/24\"}, {\"10.0.1.0/24\"})",
                    Result = "{\"10.0.0.0/24\"}"
                ],
                [
                    Description = "Preserve slow-mode input order around a split.",
                    Code = "IpSubtractSubnets({\"10.0.0.0/24\", \"192.0.2.0/24\"}, {\"10.0.0.128/25\"}, false)",
                    Result = "{\"10.0.0.0/25\", \"192.0.2.0/24\"}"
                ]
            }
        ],

    IpSubtractSubnetsDoc = Value.ReplaceType(IpSubtractSubnets, IpSubtractSubnetsType),

    /*
        Returns the IPv4 subnets common to two subnet lists.

        The VBA and JavaScript references implement intersection by taking
        the first list minus the second list, then subtracting that difference
        from the original first list. M composes the typed fast subtraction
        function twice, so duplicates and overlapping inputs are normalized by
        the same aggregation and splitting rules as IpSubtractSubnets.

        The Power Query contract replaces Excel ranges with two lists of
        subnet text. Null and empty text items are ignored as empty range
        cells; every remaining item must be text. The result is a canonical
        list of CIDR or dotted-mask subnet text produced by fast subtraction.
        Empty input on either side returns an empty list.

        Examples:
            IpCommonSubnetsDoc(
                {"10.0.0.0/24"},
                {"10.0.0.0/25"}
            ) = {"10.0.0.0/25"}
            IpCommonSubnetsDoc(
                {"10.0.0.0/24"},
                {"10.0.1.0/24"}
            ) = {}
            IpCommonSubnetsDoc(
                {"192.168.1.0/24", "192.168.2.0/24"},
                {"192.168.1.128/25"}
            ) = {"192.168.1.128/25"}
    */
    IpCommonSubnets = (firstSubnets as nullable list, secondSubnets as nullable list) as list =>
        let
            errorName = "IpCommonSubnets.InvalidInput",
            checkedFirstSubnets = IpSubnetCleanList(firstSubnets, "first", errorName),
            checkedSecondSubnets = IpSubnetCleanList(secondSubnets, "second", errorName),
            firstCandidates = IpSubnetCandidatesFromList(checkedFirstSubnets, errorName),
            secondCandidates = IpSubnetCandidatesFromList(checkedSecondSubnets, errorName),
            difference = IpSubtractSubnetCandidates(firstCandidates, secondCandidates, true),
            commonCandidates = IpSubtractSubnetCandidates(firstCandidates, difference, true)
        in
            List.Transform(commonCandidates, (candidate as record) => candidate[Output]),

    IpCommonSubnetsType =
        type function (
           firstSubnets as (type nullable list meta [
                Documentation.FieldCaption = "First IPv4 subnet list",
                Documentation.FieldDescription = "A list of IPv4 subnet text values; null and empty text items are ignored."
           ]),
           secondSubnets as (type nullable list meta [
                Documentation.FieldCaption = "Second IPv4 subnet list",
                Documentation.FieldDescription = "A list of IPv4 subnet text values; null and empty text items are ignored."
           ])
        ) as list meta [
            Documentation.Name = "IpCommonSubnets",
            Documentation.LongDescription = "Computes the intersection of two IPv4 subnet lists by composing fast IpSubtractSubnets operations. Duplicate and overlapping inputs are normalized, and the result is a canonical Power Query list. Null or empty items are ignored, non-text or malformed subnet values raise structured validation errors, and an empty input on either side returns an empty list.",
            Documentation.Examples = {
                [
                    Description = "Find the lower half shared by a /24 and a /25.",
                    Code = "IpCommonSubnets({\"10.0.0.0/24\"}, {\"10.0.0.0/25\"})",
                    Result = "{\"10.0.0.0/25\"}"
                ],
                [
                    Description = "Return an empty list for disjoint subnets.",
                    Code = "IpCommonSubnets({\"10.0.0.0/24\"}, {\"10.0.1.0/24\"})",
                    Result = "{}"
                ],
                [
                    Description = "Find an overlapping portion across a multi-subnet list.",
                    Code = "IpCommonSubnets({\"192.168.1.0/24\", \"192.168.2.0/24\"}, {\"192.168.1.128/25\"})",
                    Result = "{\"192.168.1.128/25\"}"
                ]
            }
        ],

    IpCommonSubnetsDoc = Value.ReplaceType(IpCommonSubnets, IpCommonSubnetsType),

    /*
        Looks up the most-specific matching subnet across a combined table.

        The VBA function accepts a discontiguous Excel Range and scans every
        Area and row, returning the value from a selected column for the
        longest matching subnet. Power Query has no Range.Areas value, so the
        contract accepts one already-combined table with a required Subnet
        column and a caller-supplied result column name.

        M scans the table in row order, retains the first equal-prefix match,
        and returns the selected cell as nullable text. Null and empty Subnet
        cells are ignored; remaining Subnet cells must be text. A null or
        empty result column name, missing column, malformed subnet, or invalid
        search address raises a structured validation error. No match returns
        null rather than the VBA sentinel "Not Found". The VBA baseline starts
        at 0.0.0.0/0 and requires a strictly longer prefix, so a /0 row alone
        does not count as a match; a more-specific match can supersede it.

        Examples:
            IpSubnetVLookupAreasDoc(
                "10.1.2.3",
                #table(
                    type table [Subnet = text, Name = text],
                    {{"10.0.0.0/8", "wide"}, {"10.1.0.0/16", "narrow"}}
                ),
                "Name"
            ) = "narrow"
            IpSubnetVLookupAreasDoc(
                "192.0.2.1",
                #table(type table [Subnet = text, Name = text], {{"10.0.0.0/8", "wide"}}),
                "Name"
            ) = null
    */
    IpSubnetVLookupAreas = (
        ip as nullable text,
        tableArray as nullable table,
        resultColumn as nullable text
    ) as nullable text =>
        let
            fail = IpMakeFailure("IpSubnetVLookupAreas.InvalidInput"),

            checkedTable =
                if tableArray = null then
                    fail(
                        "Lookup table cannot be null.",
                        [Input = tableArray, Component = "table", Expected = "table with a Subnet column"]
                    )
                else
                    tableArray,
            checkedResultColumn =
                if resultColumn = null or resultColumn = "" then
                    fail(
                        "Result column name cannot be null or empty.",
                        [Input = resultColumn, Component = "result column", Expected = "non-empty column name"]
                    )
                else
                    resultColumn,
            hasRequiredColumns = Table.HasColumns(checkedTable, {"Subnet", checkedResultColumn}),
            subnetValues =
                if hasRequiredColumns then
                    List.Buffer(Table.Column(checkedTable, "Subnet"))
                else
                    fail(
                        "Lookup table must contain Subnet and the selected result column.",
                        [Input = tableArray, Component = "table", Expected = "Subnet and " & checkedResultColumn & " columns"]
                    ),
            resultValues =
                if hasRequiredColumns then
                    List.Buffer(Table.Column(checkedTable, checkedResultColumn))
                else
                    {},
            searchAddressValue = IpStrToBin(ip),

            candidateFromRow = (position as number) as record =>
                let
                    rawSubnet = subnetValues{position},
                    candidate =
                        if rawSubnet = null then
                            [Matched = false, PrefixLength = -1, Value = null]
                        else if not Value.Is(rawSubnet, type text) then
                            fail(
                                "Subnet table cells must contain text.",
                                [Input = rawSubnet, Component = "Subnet", Row = position + 1, Expected = "IPv4 subnet text"]
                            )
                        else if rawSubnet = "" then
                            [Matched = false, PrefixLength = -1, Value = null]
                        else
                            let
                                parsedSubnet = IpSubnetParseWithAddressValue(rawSubnet),
                                candidateFacts = parsedSubnet[Facts],
                                searchFacts = IpSubnetFacts(searchAddressValue, candidateFacts[PrefixLength]),
                                selectedValue = Text.From(resultValues{position}, "en-US")
                            in
                                [
                                    Matched = searchFacts[NetworkKey] = candidateFacts[NetworkKey],
                                    PrefixLength = candidateFacts[PrefixLength],
                                    Value = selectedValue
                                ]
                in
                    candidate,

            bestMatch =
                List.Accumulate(
                    List.Positions(subnetValues),
                    [BestPrefix = 0, Found = false, Value = null],
                    (state as record, position as number) as record =>
                        let
                            candidate = candidateFromRow(position)
                        in
                            if candidate[Matched] and candidate[PrefixLength] > state[BestPrefix] then
                                [
                                    BestPrefix = candidate[PrefixLength],
                                    Found = true,
                                    Value = candidate[Value]
                                ]
                            else
                                state
                )
        in
            if bestMatch[Found] then
                bestMatch[Value]
            else
                null,

    IpSubnetVLookupAreasType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "Search IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           tableArray as (type nullable table meta [
                Documentation.FieldCaption = "Combined subnet table",
                Documentation.FieldDescription = "A table with a required text Subnet column and the selected result column; combine discontiguous sources before calling."
           ]),
           resultColumn as (type nullable text meta [
                Documentation.FieldCaption = "Result column",
                Documentation.FieldDescription = "The name of the table column whose value should be returned for the best match."
           ])
        ) as nullable text meta [
            Documentation.Name = "IpSubnetVLookupAreas",
            Documentation.LongDescription = "Scans a combined Power Query table with a required Subnet column and returns the selected result-column value for the most-specific matching subnet. Equal-prefix matches retain the first row. Null or empty Subnet cells are ignored; no match returns null instead of the VBA sentinel Not Found; and a /0 row alone does not count as a match because the VBA baseline requires a strictly longer prefix than its initial /0 sentinel.",
            Documentation.Examples = {
                [
                    Description = "Return the most-specific matching row's value.",
                    Code = "IpSubnetVLookupAreas(\"10.1.2.3\", #table(type table [Subnet = text, Name = text], {{\"10.0.0.0/8\", \"wide\"}, {\"10.1.0.0/16\", \"narrow\"}}), \"Name\")",
                    Result = "\"narrow\""
                ],
                [
                    Description = "Return null when no subnet matches.",
                    Code = "IpSubnetVLookupAreas(\"192.0.2.1\", #table(type table [Subnet = text, Name = text], {{\"10.0.0.0/8\", \"wide\"}}), \"Name\")",
                    Result = "null"
                ],
                [
                    Description = "Use the first row when equal-prefix matches are duplicated.",
                    Code = "IpSubnetVLookupAreas(\"10.1.2.3\", #table(type table [Subnet = text, Name = text], {{\"10.1.0.0/16\", \"first\"}, {\"10.1.0.0/16\", \"second\"}}), \"Name\")",
                    Result = "\"first\""
                ]
            }
        ],

    IpSubnetVLookupAreasDoc = Value.ReplaceType(IpSubnetVLookupAreas, IpSubnetVLookupAreasType),

    /*
        Returns the prefix length represented by a dotted-decimal IPv4 mask.

        A subnet mask is valid when it is the canonical mask for one of the
        prefix lengths 0..32. Comparing against those 33 masks is explicit,
        exact for IPv4 numbers, and rejects non-contiguous masks rather than
        assigning them an ambiguous prefix length.

        The JavaScript counterpart is ipMaskLen. Unlike its binary-string
        implementation, this M contract rejects non-contiguous masks explicitly
        instead of deriving a misleading prefix length from the last one bit.

        Examples:
            IpMaskLenDoc("255.255.255.0") = 24
            IpMaskLenDoc("0.0.0.0") = 0
            IpMaskLenDoc("255.255.255.255") = 32

        Null, malformed dotted-decimal text, and non-contiguous masks raise
        structured validation errors.
    */
    IpMaskLen = (mask as nullable text) as number =>
        let
            fail = IpMakeFailure("IpMaskLen.InvalidInput"),

            maskValue =
                if mask = null then
                    fail(
                        "IPv4 mask cannot be null.",
                        [Input = mask, Component = "mask", Expected = "dotted-decimal IPv4 mask"]
                    )
                else
                    IpStrToBin(mask),
            prefixLengths = {0..32},
            canonicalMasks =
                List.Transform(
                    prefixLengths,
                    (prefix as number) =>
                        (Number.Power(2, prefix) - 1) * Number.Power(2, 32 - prefix)
                ),
            prefixLength = List.PositionOf(canonicalMasks, maskValue)
        in
            if prefixLength >= 0 then
                prefixLength
            else
                fail(
                    "IPv4 mask must contain contiguous one bits followed by contiguous zero bits.",
                    [
                        Input = mask,
                        Component = "mask",
                        Value = maskValue,
                        Expected = "canonical IPv4 netmask"
                    ]
                ),

    IpMaskLenType =
        type function (
           mask as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 mask",
                Documentation.FieldDescription = "A dotted-decimal contiguous IPv4 subnet mask."
           ])
        ) as number meta [
            Documentation.Name = "IpMaskLen",
            Documentation.LongDescription = "Accepts canonical contiguous IPv4 subnet masks and returns a prefix length from 0 through 32.",
            Documentation.Examples = {
                [
                    Description = "Get the prefix length of a /24 mask.",
                    Code = "IpMaskLen(\"255.255.255.0\")",
                    Result = "24"
                ],
                [
                    Description = "Get the prefix length of the default route mask.",
                    Code = "IpMaskLen(\"0.0.0.0\")",
                    Result = "0"
                ],
                [
                    Description = "Get the prefix length of the host mask.",
                    Code = "IpMaskLen(\"255.255.255.255\")",
                    Result = "32"
                ]
            }
        ],

    IpMaskLenDoc = Value.ReplaceType(IpMaskLen, IpMaskLenType),

    /*
        Returns the numeric IPv4 mask represented by a subnet expression.

        The VBA implementation derives the canonical mask from the subnet
        prefix as (2^bits - 1) * 2^(32 - bits). M delegates subnet parsing and
        validation to IpSubnetLen, then performs that same bounded arithmetic
        with Number.Power. The numeric result is an intentional compatibility
        helper; public human-readable masks remain dotted text at the
        IpMask boundary.

        CIDR, dotted-mask, and unmasked address notation are supported. An
        unmasked address is treated as /32. Null, malformed addresses,
        invalid prefixes, and non-canonical masks raise structured validation
        errors through IpSubnetLen.

        Examples:
            IpMaskBinDoc("192.168.1.1/24") = 4294967040
            IpMaskBinDoc("192.168.1.0 255.255.255.0") = 4294967040
            IpMaskBinDoc("192.168.1.1") = 4294967295
            IpMaskBinDoc("192.168.1.1/0") = 0
    */
    IpMaskBin = (subnet as nullable text) as number =>
        let
            prefixLength = IpSubnetLen(subnet),
            oneBits = Number.Power(2, prefixLength) - 1,
            hostBitScale = Number.Power(2, 32 - prefixLength)
        in
            oneBits * hostBitScale,

    IpMaskBinType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as number meta [
            Documentation.Name = "IpMaskBin",
            Documentation.LongDescription = "Returns the canonical numeric IPv4 mask for a validated subnet expression using 2^prefix minus 1, shifted by the host-bit count. CIDR, dotted-decimal mask, and unmasked address notation are supported; an unmasked address is treated as /32. Invalid inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Get the numeric /24 mask from CIDR notation.",
                    Code = "IpMaskBin(\"192.168.1.1/24\")",
                    Result = "4294967040"
                ],
                [
                    Description = "Get the same numeric mask from dotted-mask notation.",
                    Code = "IpMaskBin(\"192.168.1.0 255.255.255.0\")",
                    Result = "4294967040"
                ],
                [
                    Description = "Treat an unmasked address as a /32 host subnet.",
                    Code = "IpMaskBin(\"192.168.1.1\")",
                    Result = "4294967295"
                ],
                [
                    Description = "Return the all-zero mask for /0.",
                    Code = "IpMaskBin(\"192.168.1.1/0\")",
                    Result = "0"
                ]
            }
        ],

    IpMaskBinDoc = Value.ReplaceType(IpMaskBin, IpMaskBinType),

    /*
        Returns the canonical dotted-decimal mask represented by a subnet.

        The VBA and JavaScript references accept CIDR and dotted-mask subnet
        notation and return the corresponding mask text. M composes the
        numeric IpMaskBin helper with IpBinToStr, so the subnet is parsed and
        validated once while the public result remains human-readable dotted
        IPv4 text.

        Unmasked addresses are treated as /32. Null, malformed addresses,
        invalid prefixes, and non-canonical dotted masks raise structured
        validation errors through IpMaskBin and IpSubnetLen.

        Examples:
            IpMaskDoc("192.168.1.1/24") = "255.255.255.0"
            IpMaskDoc("192.168.1.1 255.255.255.0") = "255.255.255.0"
            IpMaskDoc("192.168.1.1") = "255.255.255.255"
    */
    IpMask = (subnet as nullable text) as text =>
        let
            maskValue = IpMaskBin(subnet)
        in
            IpBinToStr(maskValue),

    IpMaskType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as text meta [
            Documentation.Name = "IpMask",
            Documentation.LongDescription = "Returns the canonical dotted-decimal mask for a validated IPv4 subnet expression. CIDR, dotted-decimal mask, and unmasked address notation are supported; an unmasked address is treated as /32. Invalid inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Get the mask from CIDR notation.",
                    Code = "IpMask(\"192.168.1.1/24\")",
                    Result = "255.255.255.0"
                ],
                [
                    Description = "Get the mask from dotted-mask notation.",
                    Code = "IpMask(\"192.168.1.1 255.255.255.0\")",
                    Result = "255.255.255.0"
                ],
                [
                    Description = "Treat an unmasked address as a /32 host subnet.",
                    Code = "IpMask(\"192.168.1.1\")",
                    Result = "255.255.255.255"
                ]
            }
        ],

    IpMaskDoc = Value.ReplaceType(IpMask, IpMaskType),

    /*
        Returns the wildcard mask represented by an IPv4 subnet.

        The VBA and JavaScript references complement the canonical subnet
        mask, so a /24 produces 0.0.0.255 and a /32 produces 0.0.0.0. M
        composes IpMaskBin with the all-ones IPv4 value and IpBinToStr; CIDR,
        dotted-mask, and unmasked address notation share the same strict
        subnet validation.

        Examples:
            IpWildMaskDoc("192.168.1.1/24") = "0.0.0.255"
            IpWildMaskDoc("192.168.1.1 255.255.255.0") = "0.0.0.255"
            IpWildMaskDoc("192.168.1.1") = "0.0.0.0"
    */
    IpWildMask = (subnet as nullable text) as text =>
        let
            maskValue = IpMaskBin(subnet),
            allBits = Number.Power(2, 32) - 1,
            wildcardValue = allBits - maskValue
        in
            IpBinToStr(wildcardValue),

    IpWildMaskType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as text meta [
            Documentation.Name = "IpWildMask",
            Documentation.LongDescription = "Returns the 32-bit complement of a validated canonical subnet mask. CIDR, dotted-decimal mask, and unmasked address notation are supported; an unmasked address is treated as /32. Invalid inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Get the wildcard mask for a /24 subnet.",
                    Code = "IpWildMask(\"192.168.1.1/24\")",
                    Result = "0.0.0.255"
                ],
                [
                    Description = "Get the wildcard mask from dotted-mask notation.",
                    Code = "IpWildMask(\"192.168.1.1 255.255.255.0\")",
                    Result = "0.0.0.255"
                ],
                [
                    Description = "Treat an unmasked address as a /32 host subnet.",
                    Code = "IpWildMask(\"192.168.1.1\")",
                    Result = "0.0.0.0"
                ]
            }
        ],

    IpWildMaskDoc = Value.ReplaceType(IpWildMask, IpWildMaskType),

    /*
        Converts an IPv4 address range into the smallest aligned CIDR list.

        The VBA and JavaScript references first expand the endpoints to an
        even first address and an odd last address, then greedily select the
        largest aligned network that does not extend beyond the remaining
        range. M expresses the former Excel 2D range result as a list of CIDR
        text and uses List.Generate for the immutable state transition.

        The optional descending parameter models ExportCellRange's ordering
        behavior even though the original IpRangeToCIDR reference does not
        expose that parameter directly. A null value means ascending order.
        A range whose normalized first address exceeds its normalized last
        address raises a structured validation error.

        Candidate prefix lengths run 0..32, a deliberate departure from the
        VBA/JavaScript references. Both references start their greedy trial
        prefix length at 1 (l = 0 then l = l + 1 before the first candidate is
        tried), so neither can ever select /0, even though their own stated
        goal is unqualified: "find the largest network which first address is
        firstAddr and which last address is not higher than lastAddr." /0
        satisfies that condition whenever the normalized range spans the
        entire address space, and every other prefix-accepting function in
        this file (IpMaskLen, IpSubnetParseWithAddressValue, IpSubnetMatch)
        already treats /0 as an ordinary, valid prefix, so excluding it here
        would be inconsistent with the rest of the library. Neither reference
        documents or tests the full-address-space case, and the JavaScript
        port shares the VBA's loop structure closely enough that their
        agreement does not corroborate the omission as intentional. Covering
        "0.0.0.0" through "255.255.255.255" therefore collapses to the single
        block "0.0.0.0/0" rather than the two /1 blocks the references would
        produce.

        Examples:
            IpRangeToCIDRDoc("10.0.0.1", "10.0.0.254")
                = {"10.0.0.0/24"}
            IpRangeToCIDRDoc("10.0.0.1", "10.0.1.63")
                = {"10.0.0.0/24", "10.0.1.0/26"}
            IpRangeToCIDRDoc("10.0.0.0", "10.0.0.255", true)
                = {"10.0.0.0/24"}
            IpRangeToCIDRDoc("0.0.0.0", "255.255.255.255")
                = {"0.0.0.0/0"}
    */
    IpRangeToCIDR = (
        firstAddress as nullable text,
        lastAddress as nullable text,
        optional descending as nullable logical
    ) as list =>
        let
            fail = IpMakeFailure("IpRangeToCIDR.InvalidInput"),

            firstValue = IpStrToBin(firstAddress),
            lastValue = IpStrToBin(lastAddress),
            normalizedFirst = Number.IntegerDivide(firstValue, 2) * 2,
            normalizedLast = Number.IntegerDivide(lastValue, 2) * 2 + 1,

            candidatePrefixes = {0..32},
            largestBlock = (current as number) as record =>
                let
                    candidates =
                        List.Transform(
                            candidatePrefixes,
                            (prefixLength as number) as record =>
                                let
                                    facts = IpSubnetFacts(current, prefixLength)
                                in
                                    [
                                        PrefixLength = facts[PrefixLength],
                                        Network = facts[NetworkKey],
                                        Broadcast = facts[NetworkKey] + facts[BlockSize] - 1
                                    ]
                        ),
                    alignedCandidates =
                        List.Select(
                            candidates,
                            (candidate as record) as logical =>
                                candidate[Network] = current
                                    and candidate[Broadcast] <= normalizedLast
                        )
                in
                    List.First(alignedCandidates),

            blocks =
                if normalizedFirst > normalizedLast then
                    fail(
                        "The normalized first address must not exceed the normalized last address.",
                        [
                            Input = [FirstAddress = firstAddress, LastAddress = lastAddress],
                            Component = "range",
                            NormalizedFirst = normalizedFirst,
                            NormalizedLast = normalizedLast,
                            Expected = "first address <= last address"
                        ]
                    )
                else
                    let
                        initialBlock = largestBlock(normalizedFirst)
                    in
                        List.Generate(
                            () => [Current = normalizedFirst] & initialBlock,
                            (state as record) as logical => state[Current] <= normalizedLast,
                            (state as record) as record =>
                                let
                                    nextStart = state[Broadcast] + 1,
                                    nextBlock =
                                        if nextStart <= normalizedLast then
                                            largestBlock(nextStart)
                                        else
                                            [PrefixLength = null, Network = null, Broadcast = null]
                                in
                                    [Current = nextStart] & nextBlock
                        ),
            ascending =
                List.Transform(
                    blocks,
                    (block as record) as text =>
                        IpBinToStr(block[Network])
                            & "/"
                            & Number.ToText(block[PrefixLength], "0", "en-US")
                ),
            descendingMode = descending ?? false
        in
            if descendingMode then
                List.Reverse(ascending)
            else
                ascending,

    IpRangeToCIDRType =
        type function (
           firstAddress as (type nullable text meta [
                Documentation.FieldCaption = "First IPv4 address",
                Documentation.FieldDescription = "The first dotted-decimal IPv4 address in the range."
           ]),
           lastAddress as (type nullable text meta [
                Documentation.FieldCaption = "Last IPv4 address",
                Documentation.FieldDescription = "The last dotted-decimal IPv4 address in the range."
           ]),
           optional descending as (type nullable logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Reverse the generated CIDR list; null means ascending order."
           ])
        ) as list meta [
            Documentation.Name = "IpRangeToCIDR",
            Documentation.LongDescription = "Expands the endpoints to the reference's even-first and odd-last boundaries, then returns the smallest aligned CIDR block list covering that normalized range, including /0 for the full address space. The result is a Power Query list of CIDR text; null descending means ascending order, and invalid or reversed normalized ranges raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Convert a single-subnet address range.",
                    Code = "IpRangeToCIDR(\"10.0.0.1\", \"10.0.0.254\")",
                    Result = "{\"10.0.0.0/24\"}"
                ],
                [
                    Description = "Convert a range requiring two CIDR blocks.",
                    Code = "IpRangeToCIDR(\"10.0.0.1\", \"10.0.1.63\")",
                    Result = "{\"10.0.0.0/24\", \"10.0.1.0/26\"}"
                ],
                [
                    Description = "Reverse the generated list.",
                    Code = "IpRangeToCIDR(\"10.0.0.0\", \"10.0.0.255\", true)",
                    Result = "{\"10.0.0.0/24\"}"
                ],
                [
                    Description = "Collapse the entire IPv4 address space to a single /0 block.",
                    Code = "IpRangeToCIDR(\"0.0.0.0\", \"255.255.255.255\")",
                    Result = "{\"0.0.0.0/0\"}"
                ]
            }
        ],

    IpRangeToCIDRDoc = Value.ReplaceType(IpRangeToCIDR, IpRangeToCIDRType),

    IpSubnetParseWithAddressValue = (
        subnet as nullable text,
        optional reason as nullable text
    ) as [Address = text, PrefixLength = number, AddressValue = number, Facts = record] =>
        let
            fail = IpMakeFailure(reason ?? "IpSubnetParse.InvalidInput"),

            parsePrefix = (prefixText as text) as number =>
                let
                    containsOnlyDigits = IpIsAsciiDigits(prefixText),
                    prefix =
                        if containsOnlyDigits then
                            Number.FromText(prefixText, "en-US")
                        else
                            null,
                    isValid =
                        if not containsOnlyDigits then
                            false
                        else
                            prefix >= 0
                                and prefix <= 32
                                and Number.RoundDown(prefix) = prefix
                in
                    if isValid then
                        prefix
                    else
                        fail(
                            "IPv4 prefix length must be an integer from 0 through 32.",
                            [
                                Input = subnet,
                                Component = "prefix length",
                                Value = prefixText,
                                Expected = "0..32"
                            ]
                        ),

            checkedSubnet =
                if subnet = null then
                    fail(
                        "IPv4 subnet cannot be null.",
                        [Input = subnet, Component = "subnet", Expected = "IPv4 subnet text"]
                    )
                else
                    subnet,
            slashPosition = Text.PositionOf(checkedSubnet, "/"),
            spacePosition = Text.PositionOf(checkedSubnet, " "),
            hasSlash = slashPosition >= 0,
            hasSpace = spacePosition >= 0,
            malformedDelimiterOrder = hasSlash and hasSpace and spacePosition < slashPosition,
            addressText =
                if malformedDelimiterOrder then
                    fail(
                        "CIDR prefix must immediately follow the IPv4 address.",
                        [Input = subnet, Component = "delimiter", Expected = "address/prefix or address mask"]
                    )
                else if hasSlash then
                    Text.Start(checkedSubnet, slashPosition)
                else if hasSpace then
                    Text.Start(checkedSubnet, spacePosition)
                else
                    checkedSubnet,
            suffixText =
                if hasSlash then
                    Text.Range(checkedSubnet, slashPosition + 1)
                else if hasSpace then
                    Text.Range(checkedSubnet, spacePosition + 1)
                else
                    null,
            prefixLength =
                if hasSlash then
                    parsePrefix(suffixText)
                else if hasSpace then
                    IpMaskLen(suffixText)
                else
                    32,
            validatedAddress = try IpStrToBin(addressText),
            addressValue =
                if validatedAddress[HasError] then
                    fail(
                        "IPv4 address must be in the dotted-decimal format.",
                        [
                            Input = subnet,
                            Component = "address",
                            Value = addressText,
                            Cause = validatedAddress[Error]
                        ]
                    )
                else
                    validatedAddress[Value],
            facts = IpSubnetFacts(addressValue, prefixLength)
        in
            [Address = addressText, PrefixLength = prefixLength, AddressValue = addressValue, Facts = facts],

    /*
        Parses an IPv4 address with an optional subnet mask.

        VBA modifies its ByRef string argument to remove the mask. M values
        are immutable, so this version returns a record containing both the
        mask-free address and its prefix length. It accepts CIDR notation
        (address/prefix) and dotted-mask notation (address mask); an address
        without a mask is treated as a /32 host address.

        Examples:
            IpSubnetParseDoc("192.168.1.1/24")
                = [Address = "192.168.1.1", PrefixLength = 24]
            IpSubnetParseDoc("192.168.1.1 255.255.255.0")
                = [Address = "192.168.1.1", PrefixLength = 24]
    */
    IpSubnetParse = (subnet as nullable text) as record =>
        let
            parsedSubnet = IpSubnetParseWithAddressValue(subnet)
        in
            [Address = parsedSubnet[Address], PrefixLength = parsedSubnet[PrefixLength]],

    IpSubnetParseType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as [Address = text, PrefixLength = number] meta [
            Documentation.Name = "IpSubnetParse",
            Documentation.LongDescription = "Returns a record with Address and PrefixLength fields. CIDR notation and dotted-decimal mask notation are supported; an address without a mask is treated as /32.",
            Documentation.Examples = {
                [
                    Description = "Parse CIDR notation.",
                    Code = "IpSubnetParse(\"192.168.1.1/24\")",
                    Result = "[Address = \"192.168.1.1\", PrefixLength = 24]"
                ],
                [
                    Description = "Parse dotted-mask notation.",
                    Code = "IpSubnetParse(\"192.168.1.1 255.255.255.0\")",
                    Result = "[Address = \"192.168.1.1\", PrefixLength = 24]"
                ],
                [
                    Description = "Treat an address without a mask as a /32 host address.",
                    Code = "IpSubnetParse(\"192.168.1.1\")",
                    Result = "[Address = \"192.168.1.1\", PrefixLength = 32]"
                ]
            }
        ],

    IpSubnetParseDoc = Value.ReplaceType(IpSubnetParse, IpSubnetParseType),

    /*
        Converts an IPv4 subnet to its numeric network address.

        The VBA implementation parses the subnet prefix, clears the host
        bits in each octet, and accumulates the resulting bytes in base 256.
        M can express the same projection on the validated IPv4 integer: a
        block size of 2^(32 - prefix) identifies the host portion, and integer
        division followed by multiplication clears it. All values remain in
        the exact IPv4 range 0..4294967295.

        No direct JavaScript public counterpart was identified. Related
        JavaScript IpNet construction normalizes the address to the same
        network boundary, but this function returns the numeric key rather
        than an object or formatted address.

        Examples:
            IpSubnetToBinDoc("1.2.3.4/24") = 16909056
            IpSubnetToBinDoc("1.2.3.0/24") = 16909056
            IpSubnetToBinDoc("1.2.3.4/32") = 16909060
            IpSubnetToBinDoc("1.2.3.4/0") = 0

        Null, malformed, invalid-prefix, and non-canonical-mask inputs raise
        the structured validation errors from IpSubnetParse.
    */
    IpSubnetToBin = (subnet as nullable text) as number =>
        let
            parsedSubnet = IpSubnetParseWithAddressValue(subnet)
        in
            parsedSubnet[Facts][NetworkKey],

    IpSubnetToBinType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as number meta [
            Documentation.Name = "IpSubnetToBin",
            Documentation.LongDescription = "Returns the IPv4 integer for the subnet's network address in the range 0 through 4294967295. Host bits are cleared according to a CIDR or canonical dotted-decimal mask; an unmasked address is treated as /32. Invalid or null inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Clear the host byte from a /24 subnet.",
                    Code = "IpSubnetToBin(\"1.2.3.4/24\")",
                    Result = "16909056"
                ],
                [
                    Description = "A network address already has the same numeric key.",
                    Code = "IpSubnetToBin(\"1.2.3.0/24\")",
                    Result = "16909056"
                ],
                [
                    Description = "A /32 preserves every address bit.",
                    Code = "IpSubnetToBin(\"1.2.3.4/32\")",
                    Result = "16909060"
                ],
                [
                    Description = "A /0 subnet has the all-zero network key.",
                    Code = "IpSubnetToBin(\"1.2.3.4/0\")",
                    Result = "0"
                ]
            }
        ],

    IpSubnetToBinDoc = Value.ReplaceType(IpSubnetToBin, IpSubnetToBinType),

    /*
        Finds the best matching IPv4 subnet in a table.

        The VBA and JavaScript references accept an Excel range or an array
        whose first column contains subnet expressions and return a 1-based
        row number, or 0 when nothing matches. M has no Excel Range value, so
        this contract requires a table with a text column named Subnet;
        additional columns are allowed and are not inspected.

        Best-match mode scans the table in its original order and retains the
        most specific matching prefix. Equal-length matches retain the first
        row, matching the reference's strict greater-than update. A /0 row is
        deliberately excluded, aligning this function with the VBA-compatible
        IpSubnetVLookupAreas baseline. Fast mode validates, then buffers parsed
        candidates for stable indexed access and binary search. Its rows must
        be non-empty, sorted by ascending network address, and non-overlapping;
        invalid fast-mode input raises a structured error instead of returning
        a potentially incorrect row.

        Examples:
            IpSubnetMatchDoc(
                "10.1.2.3",
                #table(type table [Subnet = text], {{"10.0.0.0/8"}, {"10.1.0.0/16"}})
            ) = 2
            IpSubnetMatchDoc(
                "192.0.2.1",
                #table(type table [Subnet = text], {{"10.0.0.0/8"}})
            ) = 0

        Null or empty Subnet cells are ignored in best-match mode. Null tables,
        missing Subnet columns, non-text subnet cells, malformed addresses,
        invalid prefixes, and non-canonical masks raise structured
        validation errors. A null fast value means false.
    */
    IpSubnetMatch = (
        ip as nullable text,
        tableArray as nullable table,
        optional fast as nullable logical
    ) as number =>
        let
            fail = IpMakeFailure("IpSubnetMatch.InvalidInput"),

            checkedTable =
                if tableArray = null then
                    fail(
                        "Subnet table cannot be null.",
                        [Input = tableArray, Component = "table", Expected = "table with a Subnet column"]
                    )
                else
                    tableArray,
            subnetValues =
                if Table.HasColumns(checkedTable, "Subnet") then
                    Table.Column(checkedTable, "Subnet")
                else
                    fail(
                        "Subnet table must contain a Subnet column.",
                        [Input = tableArray, Component = "table", Expected = "Subnet column"]
                    ),
            fastMode = fast ?? false,
            searchSubnet = IpSubnetParseWithAddressValue(ip),
            searchAddressValue = searchSubnet[AddressValue],
            searchPrefixLength = searchSubnet[PrefixLength],
            searchNetworkKey = searchSubnet[Facts][NetworkKey],

            candidateFromValue = (rawSubnet as any, rowNumber as number) as nullable record =>
                if rawSubnet = null or (Value.Is(rawSubnet, type text) and rawSubnet = "") then
                    null
                else
                    let
                        checkedSubnet =
                            if not Value.Is(rawSubnet, type text) then
                                fail(
                                    "Subnet table cells must contain text.",
                                    [Input = rawSubnet, Component = "Subnet", Row = rowNumber, Expected = "IPv4 subnet text"]
                                )
                            else
                                rawSubnet,
                        parsedSubnet = IpSubnetParseWithAddressValue(checkedSubnet),
                        facts = parsedSubnet[Facts]
                    in
                        [
                            Row = rowNumber,
                            Subnet = checkedSubnet,
                            PrefixLength = facts[PrefixLength],
                            NetworkKey = facts[NetworkKey],
                            BlockSize = facts[BlockSize]
                        ],

            candidateMatches = (candidate as record) as logical =>
                if searchPrefixLength < candidate[PrefixLength] then
                    false
                else
                    IpSubnetFacts(searchAddressValue, candidate[PrefixLength])[NetworkKey] = candidate[NetworkKey],

            bestMatch =
                List.Accumulate(
                    List.Positions(subnetValues),
                    [BestPrefix = 0, Row = 0],
                    (state as record, position as number) as record =>
                        let
                            candidate = candidateFromValue(subnetValues{position}, position + 1),
                            matches = if candidate = null then false else candidateMatches(candidate)
                        in
                            if matches and candidate[PrefixLength] > state[BestPrefix] then
                                [BestPrefix = candidate[PrefixLength], Row = candidate[Row]]
                            else
                                state
                ),

            fastCandidates =
                if fastMode then
                    List.Buffer(
                        List.Transform(
                            List.Positions(subnetValues),
                            (position as number) as record =>
                                let
                                    candidate = candidateFromValue(subnetValues{position}, position + 1)
                                in
                                    if candidate = null then
                                        fail(
                                            "Fast mode requires non-empty Subnet cells.",
                                            [Input = subnetValues{position}, Component = "Subnet", Row = position + 1, Expected = "sorted non-empty IPv4 subnet text"]
                                        )
                                    else
                                        candidate
                        )
                    )
                else
                    {},
            fastPrecondition =
                if not fastMode then
                    true
                else
                    List.Accumulate(
                        fastCandidates,
                        [Previous = null, Valid = true],
                        (state as record, candidate as record) as record =>
                            let
                                previous = state[Previous],
                                orderedAndDisjoint =
                                    if previous = null then
                                        true
                                    else
                                        previous[NetworkKey] + previous[BlockSize] <= candidate[NetworkKey]
                            in
                                if state[Valid] and orderedAndDisjoint then
                                    [Previous = candidate, Valid = true]
                                else
                                    [Previous = candidate, Valid = false]
                    )[Valid],
            validatedFastCandidates =
                if fastPrecondition then
                    fastCandidates
                else
                    fail(
                        "Fast mode requires subnets sorted by network address without overlaps.",
                        [Input = tableArray, Component = "Subnet", Expected = "ascending, non-overlapping, non-empty IPv4 subnet rows"]
                    ),
            rowCount = List.Count(validatedFastCandidates),
            fastMatch =
                if rowCount = 0 then
                    0
                else
                    let
                        binarySearch = (low as number, high as number) as record =>
                            if low >= high then
                                [Index = low, Candidate = validatedFastCandidates{low}]
                            else
                                let
                                    midpoint = Number.IntegerDivide(low + high + 1, 2),
                                    candidate = validatedFastCandidates{midpoint}
                                in
                                    if searchNetworkKey < candidate[NetworkKey] then
                                        @binarySearch(low, midpoint - 1)
                                    else
                                        @binarySearch(midpoint, high),
                        selected = binarySearch(0, rowCount - 1),
                        selectedCandidate = selected[Candidate]
                    in
                        if selectedCandidate[PrefixLength] > 0 and candidateMatches(selectedCandidate) then
                            selectedCandidate[Row]
                        else
                            0
        in
            if fastMode then
                fastMatch
            else
                bestMatch[Row],

    IpSubnetMatchType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "Search IPv4 address or subnet",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with optional CIDR or dotted-mask notation."
           ]),
           tableArray as (type nullable table meta [
                Documentation.FieldCaption = "Subnet table",
                Documentation.FieldDescription = "A table with a required text column named Subnet; additional columns are permitted."
           ]),
           optional fast as (type nullable logical meta [
                Documentation.FieldCaption = "Fast mode",
                Documentation.FieldDescription = "Use binary search; the Subnet rows must be sorted by network address and non-overlapping."
           ])
        ) as number meta [
            Documentation.Name = "IpSubnetMatch",
            Documentation.LongDescription = "Searches the required Subnet column of a Power Query table. Best-match mode accepts any row order, ignores null or empty Subnet cells, and returns the most specific matching subnet; fast mode uses binary search and requires ascending, non-overlapping, non-empty subnet rows. Returns 0 when no subnet matches. Null or malformed inputs and other invalid table cells raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Return the most specific matching row using best-match mode.",
                    Code = "IpSubnetMatch(\"10.1.2.3\", #table(type table [Subnet = text], {{\"10.0.0.0/8\"}, {\"10.1.0.0/16\"}}))",
                    Result = "2"
                ],
                [
                    Description = "Return zero when no subnet contains the search address.",
                    Code = "IpSubnetMatch(\"192.0.2.1\", #table(type table [Subnet = text], {{\"10.0.0.0/8\"}}))",
                    Result = "0"
                ],
                [
                    Description = "Use fast mode with sorted, non-overlapping rows.",
                    Code = "IpSubnetMatch(\"10.1.2.3\", #table(type table [Subnet = text], {{\"10.0.0.0/16\"}, {\"10.1.0.0/16\"}}), true)",
                    Result = "2"
                ]
            }
        ],

    IpSubnetMatchDoc = Value.ReplaceType(IpSubnetMatch, IpSubnetMatchType),

    /*
        Projects the selected value from the most-specific containing subnet.

        The VBA and JavaScript references expose this operation as a
        VLOOKUP-style scalar function over a Range or array. M has no Excel
        Range concept, so the public contract is a table with a required
        Subnet column and a caller-selected result column. The result is
        returned as nullable text to preserve the VBA function's String
        boundary. The VBA reference returns the sentinel string "Not Found"
        for no match, which this M port deliberately avoids in favor of null;
        as a consequence, null means either no subnet matched or a subnet
        matched but the selected result column held null for that row -- this
        function does not distinguish the two, the same way a SQL left join's
        null result column does not. Call IpSubnetMatch directly (it returns
        0 for no match and a real row number otherwise) when a caller needs
        to tell those cases apart.

        A normal Table.Join is not sufficient here: equality joins do not
        express IPv4 containment or longest-prefix selection. M therefore
        delegates subnet matching to IpSubnetMatch, then projects the
        selected row from the requested table column. Best-match mode accepts
        arbitrary row order; fast mode retains IpSubnetMatch's requirement
        for ascending, non-overlapping subnet rows.

        Examples:
            IpSubnetJoinDoc(
                "10.1.2.3",
                #table(
                    type table [Subnet = text, Name = text],
                    {{"10.0.0.0/8", "wide"}, {"10.1.0.0/16", "narrow"}}
                ),
                "Name"
            ) = "narrow"
            IpSubnetJoinDoc(
                "192.0.2.1",
                #table(type table [Subnet = text, Name = text], {{"10.0.0.0/8", "wide"}}),
                "Name"
            ) = null
    */
    IpSubnetJoin = (
        ip as nullable text,
        tableArray as nullable table,
        resultColumn as nullable text,
        optional fast as nullable logical
    ) as nullable text =>
        let
            fail = IpMakeFailure("IpSubnetJoin.InvalidInput"),

            checkedTable =
                if tableArray = null then
                    fail(
                        "Subnet table cannot be null.",
                        [Input = tableArray, Component = "table", Expected = "table with Subnet and result columns"]
                    )
                else
                    tableArray,
            checkedResultColumn =
                if resultColumn = null or resultColumn = "" then
                    fail(
                        "Result column name cannot be null or empty.",
                        [Input = resultColumn, Component = "result column", Expected = "non-empty column name"]
                    )
                else
                    resultColumn,
            hasRequiredColumns = Table.HasColumns(checkedTable, {"Subnet", checkedResultColumn}),
            validatedTable =
                if hasRequiredColumns then
                    checkedTable
                else
                    fail(
                        "Subnet table must contain Subnet and the selected result column.",
                        [Input = tableArray, Component = "table", Expected = "Subnet and " & checkedResultColumn & " columns"]
                    ),
            bufferedTable = Table.Buffer(validatedTable),
            resultValues = Table.Column(bufferedTable, checkedResultColumn),
            matchedRow = IpSubnetMatch(ip, bufferedTable, fast)
        in
            if matchedRow = 0 then
                null
            else
                Text.From(resultValues{matchedRow - 1}, "en-US"),

    IpSubnetJoinType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "Search IPv4 address or subnet",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with optional CIDR or dotted-mask notation."
           ]),
           tableArray as (type nullable table meta [
                Documentation.FieldCaption = "Subnet table",
                Documentation.FieldDescription = "A table with required Subnet and selected result columns."
           ]),
           resultColumn as (type nullable text meta [
                Documentation.FieldCaption = "Result column",
                Documentation.FieldDescription = "The table column whose value should be returned for the most-specific match."
           ]),
           optional fast as (type nullable logical meta [
                Documentation.FieldCaption = "Fast mode",
                Documentation.FieldDescription = "Use binary search; Subnet rows must be sorted by network address and non-overlapping."
           ])
        ) as nullable text meta [
            Documentation.Name = "IpSubnetJoin",
            Documentation.LongDescription = "Matches an IPv4 address or subnet against a Power Query table's Subnet column and returns the selected result-column value from the most-specific containing subnet. Best-match mode accepts arbitrary row order; fast mode uses binary search and requires ascending, non-overlapping rows. No match returns null rather than the VBA sentinel Not Found, and a matched row whose result-column value is itself null also returns null; these two cases are not distinguished. Call IpSubnetMatch directly to tell them apart. Matched values are converted to text to preserve the reference String boundary.",
            Documentation.Examples = {
                [
                    Description = "Return the value from the most-specific matching subnet.",
                    Code = "IpSubnetJoin(\"10.1.2.3\", #table(type table [Subnet = text, Name = text], {{\"10.0.0.0/8\", \"wide\"}, {\"10.1.0.0/16\", \"narrow\"}}), \"Name\")",
                    Result = "\"narrow\""
                ],
                [
                    Description = "Return null when no subnet contains the search value.",
                    Code = "IpSubnetJoin(\"192.0.2.1\", #table(type table [Subnet = text, Name = text], {{\"10.0.0.0/8\", \"wide\"}}), \"Name\")",
                    Result = "null"
                ],
                [
                    Description = "Use fast mode with sorted, non-overlapping subnet rows.",
                    Code = "IpSubnetJoin(\"10.1.2.3\", #table(type table [Subnet = text, Name = text], {{\"10.0.0.0/16\", \"wide\"}, {\"10.1.0.0/16\", \"narrow\"}}), \"Name\", true)",
                    Result = "\"narrow\""
                ]
            }
        ],

    IpSubnetJoinDoc = Value.ReplaceType(IpSubnetJoin, IpSubnetJoinType),

    /*
        Removes an optional IPv4 subnet mask from a subnet expression.

        The executable VBA implementation returns the text before the first
        slash, or before the first space when no slash is present. Its two
        IPv4 comments incorrectly show an address ending in .0 returning .1;
        M follows the executable body and preserves the address text. The
        JavaScript counterpart constructs an IpNet and therefore normalizes
        host bits before returning its address. This M contract intentionally
        follows the VBA operation and keeps host bits unchanged.

        Delegating to IpSubnetParse makes the IPv4 contract explicit: null,
        malformed addresses, invalid CIDR prefixes, and non-canonical dotted
        masks raise structured validation errors. The returned Address field
        is then selected without parsing the subnet a second time.

        Examples:
            IpWithoutMaskDoc("192.168.1.0/24") = "192.168.1.0"
            IpWithoutMaskDoc("192.168.1.193/24") = "192.168.1.193"
            IpWithoutMaskDoc("192.168.1.0 255.255.255.0") = "192.168.1.0"
            IpWithoutMaskDoc("192.168.1.1") = "192.168.1.1"
    */
    IpWithoutMask = (subnet as nullable text) as text =>
        let
            parsedSubnet = IpSubnetParse(subnet)
        in
            parsedSubnet[Address],

    IpWithoutMaskType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as text meta [
            Documentation.Name = "IpWithoutMask",
            Documentation.LongDescription = "Returns the validated dotted-decimal IPv4 address before an optional CIDR or dotted-decimal mask. Host bits are preserved; unlike the JavaScript IpNet implementation, this function does not normalize the address to a network boundary. Invalid or null inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Remove CIDR notation from a network address.",
                    Code = "IpWithoutMask(\"192.168.1.0/24\")",
                    Result = "\"192.168.1.0\""
                ],
                [
                    Description = "Preserve host bits while removing CIDR notation.",
                    Code = "IpWithoutMask(\"192.168.1.193/24\")",
                    Result = "\"192.168.1.193\""
                ],
                [
                    Description = "Remove dotted-decimal mask notation.",
                    Code = "IpWithoutMask(\"192.168.1.0 255.255.255.0\")",
                    Result = "\"192.168.1.0\""
                ],
                [
                    Description = "Leave an already unmasked IPv4 address unchanged.",
                    Code = "IpWithoutMask(\"192.168.1.1\")",
                    Result = "\"192.168.1.1\""
                ]
            }
        ],

    IpWithoutMaskDoc = Value.ReplaceType(IpWithoutMask, IpWithoutMaskType),

    /*
        Clears the host bits from an IPv4 subnet expression.

        The VBA implementation removes the mask, applies the canonical mask
        with IpAnd, and appends the original suffix with its original text.
        M reproduces that projection with one parsed subnet and safe numeric
        IPv4 arithmetic: the address becomes the network boundary, while the
        first CIDR or space delimiter and all following characters are kept
        verbatim. An unmasked address therefore remains unmasked.

        This agrees with JavaScript for canonical CIDR and dotted-mask input.
        Unlike JavaScript IpNet formatting, M deliberately preserves source
        suffix spelling such as a leading-zero CIDR prefix or extra spacing,
        matching the VBA Mid(net, Len(ip) + 1) behavior.

        Null, malformed addresses, invalid prefixes, and non-canonical dotted
        masks raise structured validation errors through IpSubnetParse.

        Examples:
            IpClearHostBitsDoc("192.168.1.1/24") = "192.168.1.0/24"
            IpClearHostBitsDoc("192.168.1.193 255.255.255.128")
                = "192.168.1.128 255.255.255.128"
            IpClearHostBitsDoc("192.168.1.193/024") = "192.168.1.0/024"
    */
    IpClearHostBits = (subnet as nullable text) as text =>
        let
            parsedSubnet = IpSubnetParseWithAddressValue(subnet),
            networkValue = parsedSubnet[Facts][NetworkKey],
            slashPosition = Text.PositionOf(subnet, "/"),
            spacePosition = Text.PositionOf(subnet, " "),
            suffixPosition =
                if slashPosition >= 0 then
                    slashPosition
                else
                    spacePosition,
            suffix =
                if suffixPosition < 0 then
                    ""
                else
                    Text.Range(subnet, suffixPosition)
        in
            IpBinToStr(networkValue) & suffix,

    IpClearHostBitsType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-decimal mask."
           ])
        ) as text meta [
            Documentation.Name = "IpClearHostBits",
            Documentation.LongDescription = "Returns the network-boundary IPv4 address with the original CIDR or dotted-mask suffix text preserved. Unmasked addresses remain unmasked; null, malformed, invalid-prefix, and non-canonical-mask inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Clear host bits from CIDR notation.",
                    Code = "IpClearHostBits(\"192.168.1.1/24\")",
                    Result = "\"192.168.1.0/24\""
                ],
                [
                    Description = "Clear host bits from dotted-mask notation.",
                    Code = "IpClearHostBits(\"192.168.1.193 255.255.255.128\")",
                    Result = "\"192.168.1.128 255.255.255.128\""
                ],
                [
                    Description = "Preserve original prefix spelling.",
                    Code = "IpClearHostBits(\"192.168.1.193/024\")",
                    Result = "\"192.168.1.0/024\""
                ]
            }
        ],

    IpClearHostBitsDoc = Value.ReplaceType(IpClearHostBits, IpClearHostBitsType),

    /*
        Returns true when an IPv4 address is in a subnet.

        The VBA implementation parses the subnet to obtain its prefix length
        and compares that many leading bits of the address and subnet. This
        M version returns the same logical result while keeping the parsed
        subnet state in a record, because M has no ByRef mutation.

        The JavaScript counterpart is ipIsInSubnet, whose IpNet.matchIp
        comparison has the same prefix semantics. Public inputs remain
        dotted-decimal text; IpComp performs the safe IPv4 numeric comparison
        after IpSubnetParse validates the subnet.

        Examples:
            IpIsInSubnetDoc("192.168.1.35", "192.168.1.32/29") = true
            IpIsInSubnetDoc("192.168.1.35", "192.168.1.32 255.255.255.248") = true
            IpIsInSubnetDoc("192.168.1.41", "192.168.1.32/29") = false

        Null, malformed, out-of-range, and non-canonical-mask inputs raise
        structured validation errors from IpStrToBin or IpSubnetParse.
    */
    IpIsInSubnet = (ip as nullable text, subnet as nullable text) as logical =>
        let
            parsedSubnet = IpSubnetParse(subnet)
        in
            IpComp(ip, parsedSubnet[Address], parsedSubnet[PrefixLength]),

    IpIsInSubnetType =
        type function (
           ip as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 address",
                Documentation.FieldDescription = "A dotted-decimal IPv4 address with four octets."
           ]),
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as logical meta [
            Documentation.Name = "IpIsInSubnet",
            Documentation.LongDescription = "Returns true when the validated IPv4 address shares the subnet's prefix. Accepts CIDR, dotted-decimal mask, and unmasked subnet notation; an unmasked subnet is treated as /32. Invalid inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "An address inside a /29 subnet.",
                    Code = "IpIsInSubnet(\"192.168.1.35\", \"192.168.1.32/29\")",
                    Result = "true"
                ],
                [
                    Description = "Dotted-decimal mask notation is supported.",
                    Code = "IpIsInSubnet(\"192.168.1.35\", \"192.168.1.32 255.255.255.248\")",
                    Result = "true"
                ],
                [
                    Description = "An address outside the /29 subnet.",
                    Code = "IpIsInSubnet(\"192.168.1.41\", \"192.168.1.32/29\")",
                    Result = "false"
                ]
            }
        ],

    IpIsInSubnetDoc = Value.ReplaceType(IpIsInSubnet, IpIsInSubnetType),

    /*
        Returns true when subnet1 is contained by subnet2.

        The VBA implementation compares subnet1's address with subnet2's
        prefix after rejecting a candidate whose prefix is shorter than the
        containing subnet's prefix. The JavaScript counterpart applies the
        same rule through IpNet.matchSubnet. M values are immutable, so both
        expressions are parsed into records and their address/prefix fields
        are reused explicitly.

        This function preserves the parser's address text rather than
        normalizing host bits. Containment is determined by the supplied
        address and the containing subnet's prefix, which matches the
        reference behavior for inputs such as a host address written with a
        subnet suffix.

        Examples:
            IpSubnetIsInSubnetDoc("192.168.1.35/30", "192.168.1.32/29") = true
            IpSubnetIsInSubnetDoc("192.168.1.41/30", "192.168.1.32/29") = false
            IpSubnetIsInSubnetDoc("192.168.1.35/28", "192.168.1.32/29") = false
            IpSubnetIsInSubnetDoc("192.168.0.128 255.255.255.128", "192.168.0.0 255.255.255.0") = true

        Null, malformed, out-of-range, and non-canonical-mask inputs raise
        the structured validation errors from IpSubnetParse.
    */
    IpSubnetIsInSubnet = (subnet1 as nullable text, subnet2 as nullable text) as logical =>
        let
            firstSubnet = IpSubnetParse(subnet1),
            containingSubnet = IpSubnetParse(subnet2),
            firstPrefixLength = firstSubnet[PrefixLength],
            containingPrefixLength = containingSubnet[PrefixLength]
        in
            if firstPrefixLength < containingPrefixLength then
                false
            else
                IpComp(
                    firstSubnet[Address],
                    containingSubnet[Address],
                    containingPrefixLength
                ),

    IpSubnetIsInSubnetType =
        type function (
           subnet1 as (type nullable text meta [
                Documentation.FieldCaption = "Candidate IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ]),
           subnet2 as (type nullable text meta [
                Documentation.FieldCaption = "Containing IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as logical meta [
            Documentation.Name = "IpSubnetIsInSubnet",
            Documentation.LongDescription = "Returns true when the candidate subnet has at least as many prefix bits as the containing subnet and its address shares the containing subnet's prefix. Accepts CIDR, dotted-decimal mask, and unmasked IPv4 notation. Invalid inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "A /30 address range inside a /29 subnet.",
                    Code = "IpSubnetIsInSubnet(\"192.168.1.35/30\", \"192.168.1.32/29\")",
                    Result = "true"
                ],
                [
                    Description = "An address outside the containing /29 subnet.",
                    Code = "IpSubnetIsInSubnet(\"192.168.1.41/30\", \"192.168.1.32/29\")",
                    Result = "false"
                ],
                [
                    Description = "A candidate prefix shorter than the containing prefix is not contained.",
                    Code = "IpSubnetIsInSubnet(\"192.168.1.35/28\", \"192.168.1.32/29\")",
                    Result = "false"
                ],
                [
                    Description = "Dotted-decimal mask notation is supported.",
                    Code = "IpSubnetIsInSubnet(\"192.168.0.128 255.255.255.128\", \"192.168.0.0 255.255.255.0\")",
                    Result = "true"
                ]
            }
        ],

    IpSubnetIsInSubnetDoc = Value.ReplaceType(IpSubnetIsInSubnet, IpSubnetIsInSubnetType),

    /*
        Returns the prefix length from an IPv4 subnet expression.

        The VBA implementation finds a CIDR prefix after "/", derives a
        prefix from a dotted-decimal mask after a space, and treats an
        unmasked address as a /32 host address. M values are immutable, so
        this function delegates to IpSubnetParse and selects its
        PrefixLength field rather than re-parsing the subnet independently.

        The JavaScript counterpart is ipSubnetLen, which reads the len field
        from IpNet. This M contract preserves the documented results while
        explicitly rejecting null, malformed addresses, invalid prefix
        lengths, and non-canonical masks through IpSubnetParse.

        Examples:
            IpSubnetLenDoc("192.168.1.1/24") = 24
            IpSubnetLenDoc("192.168.1.1 255.255.255.0") = 24
            IpSubnetLenDoc("192.168.1.1") = 32
    */
    IpSubnetLen = (subnet as nullable text) as number =>
        let
            parsedSubnet = IpSubnetParse(subnet)
        in
            parsedSubnet[PrefixLength],

    IpSubnetLenType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as number meta [
            Documentation.Name = "IpSubnetLen",
            Documentation.LongDescription = "Accepts a dotted IPv4 address, CIDR notation, or dotted-decimal mask notation and returns a prefix length from 0 through 32. An unmasked address is treated as /32; null, malformed addresses, invalid prefixes, and non-canonical masks raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Get the prefix length from CIDR notation.",
                    Code = "IpSubnetLen(\"192.168.1.1/24\")",
                    Result = "24"
                ],
                [
                    Description = "Get the prefix length from dotted-mask notation.",
                    Code = "IpSubnetLen(\"192.168.1.1 255.255.255.0\")",
                    Result = "24"
                ],
                [
                    Description = "Treat an address without a mask as a /32 host address.",
                    Code = "IpSubnetLen(\"192.168.1.1\")",
                    Result = "32"
                ]
            }
        ],

    IpSubnetLenDoc = Value.ReplaceType(IpSubnetLen, IpSubnetLenType),

    /*
        Sorts and summarizes IPv4 routes without merging different next hops.

        The VBA and JavaScript references normalize each route, sort by
        network and prefix length, remove duplicate routes, remove a route
        contained by a larger route only when the next hop is equal, and join
        aligned equal-size routes with that same next hop. M represents each
        route as a record, aggregates each exact next-hop group with immutable
        list operations, and sorts the combined result back into network order.
        Grouping makes same-next-hop containment transitive; this deliberately
        preserves the documented semantic result rather than an
        order-sensitive nested route that the mutable reverse scan can leave
        behind after unrelated routes are removed. Each sorted group is folded
        once onto an immutable linked stack; merge collapse recurses through at
        most the 32 IPv4 prefix levels instead of rebuilding the whole list.

        The Power Query contract replaces the Excel single-column Range with a
        nullable list of route text. IpParseRoute supplies the subnet and
        delimiter-leading next-hop suffix once; CIDR, dotted-mask, and
        unmasked host notation are accepted. Empty cells are ignored, while
        non-text or malformed routes raise structured validation errors. A
        null descending value means ascending order.

        Examples:
            IpRouteAggregateArrayDoc({
                "10.0.0.0/25 gw",
                "10.0.0.128/25 gw"
            }) = {"10.0.0.0/24 gw"}
            IpRouteAggregateArrayDoc({
                "10.0.0.0/25 gw-a",
                "10.0.0.128/25 gw-b"
            }) = {"10.0.0.0/25 gw-a", "10.0.0.128/25 gw-b"}
            IpRouteAggregateArrayDoc({
                "192.0.2.0 255.255.255.128 next-hop",
                "192.0.2.128 255.255.255.128 next-hop"
            }, true) = {"192.0.2.0 255.255.255.0 next-hop"}
    */
    IpRouteAggregateArray = (
        routes as nullable list,
        optional descending as nullable logical
    ) as list =>
        let
            fail = IpMakeFailure("IpRouteAggregateArray.InvalidInput"),

            checkedRoutes =
                if routes = null then
                    fail(
                        "Route list cannot be null.",
                        [Input = routes, Component = "routes", Expected = "list of route text"]
                    )
                else
                    routes,
            nonEmptyRoutes =
                List.Select(
                    checkedRoutes,
                    (rawRoute as any) as logical =>
                        if rawRoute = null then
                            false
                        else if Value.Is(rawRoute, type text) then
                            Text.Trim(rawRoute) <> ""
                        else
                            true
                ),
            bufferedRoutes = List.Buffer(nonEmptyRoutes),

            routeRecord = (rawRoute as any, originalIndex as number) as record =>
                let
                    checkedRoute =
                        if not Value.Is(rawRoute, type text) then
                            fail(
                                "Route list items must contain text.",
                                [Input = rawRoute, Component = "route", Expected = "IPv4 route text"]
                            )
                        else
                            rawRoute,
                    parsedRoute = IpParseRoute(checkedRoute),
                    subnetText = parsedRoute[Subnet],
                    parsedSubnet = IpSubnetParseWithAddressValue(subnetText),
                    facts = parsedSubnet[Facts],
                    usesCidr =
                        Text.PositionOf(subnetText, "/") >= 0
                            or Text.PositionOf(subnetText, " ") < 0
                in
                    [
                        NetworkKey = facts[NetworkKey],
                        PrefixLength = facts[PrefixLength],
                        UsesCidr = usesCidr,
                        NextHop = parsedRoute[NextHop],
                        FilteredIndex = originalIndex
                    ],

            routeRecords =
                List.Transform(
                    List.Positions(bufferedRoutes),
                    (position as number) => routeRecord(bufferedRoutes{position}, position)
                ),
            bufferedRouteRecords = List.Buffer(routeRecords),
            compareRoutes = (left as record, right as record) as number =>
                if left[NetworkKey] < right[NetworkKey] then
                    -1
                else if left[NetworkKey] > right[NetworkKey] then
                    1
                else if left[PrefixLength] < right[PrefixLength] then
                    -1
                else if left[PrefixLength] > right[PrefixLength] then
                    1
                else if left[FilteredIndex] < right[FilteredIndex] then
                    -1
                else if left[FilteredIndex] > right[FilteredIndex] then
                    1
                else
                    0,

            containsRoute = (outer as record, inner as record) as logical =>
                inner[PrefixLength] >= outer[PrefixLength]
                    and IpSubnetFacts(inner[NetworkKey], outer[PrefixLength])[NetworkKey] = outer[NetworkKey],

            joinedRoute = (left as record, right as record) as nullable record =>
                if left[PrefixLength] <> right[PrefixLength] or left[PrefixLength] = 0 then
                    null
                else
                    let
                        parentPrefixLength = left[PrefixLength] - 1,
                        leftParentFacts = IpSubnetFacts(left[NetworkKey], parentPrefixLength),
                        rightParentFacts = IpSubnetFacts(right[NetworkKey], parentPrefixLength)
                    in
                        if leftParentFacts[NetworkKey] = rightParentFacts[NetworkKey] then
                            [
                                NetworkKey = leftParentFacts[NetworkKey],
                                PrefixLength = parentPrefixLength,
                                UsesCidr = left[UsesCidr],
                                NextHop = left[NextHop],
                                FilteredIndex = left[FilteredIndex]
                            ]
                        else
                            null,

            pushAggregate = (stack as nullable record, route as record) as record =>
                if stack = null then
                    [Value = route, Rest = null]
                else
                    let
                        previous = stack[Value],
                        containsCurrent = containsRoute(previous, route),
                        merged = if containsCurrent then null else joinedRoute(previous, route)
                    in
                        if containsCurrent then
                            stack
                        else if merged <> null then
                            @pushAggregate(stack[Rest], merged)
                        else
                            [Value = route, Rest = stack],

            aggregateGroup = (items as list) as list =>
                let
                    stack =
                        List.Accumulate(
                            items,
                            null,
                            (state as nullable record, route as record) => pushAggregate(state, route)
                        ),
                    reverseResult =
                        List.Generate(
                            () => stack,
                            each _ <> null,
                            each _[Rest],
                            each _[Value]
                        )
                in
                    List.Reverse(reverseResult),

            renderRoute = (route as record) as text =>
                let
                    facts = IpSubnetFacts(route[NetworkKey], route[PrefixLength]),
                    networkText = IpBinToStr(facts[NetworkKey]),
                    maskValue = (Number.Power(2, route[PrefixLength]) - 1) * facts[BlockSize],
                    subnetText =
                        if route[UsesCidr] then
                            networkText & "/" & Number.ToText(route[PrefixLength], "0", "en-US")
                        else
                            networkText & " " & IpBinToStr(maskValue)
                in
                    subnetText & route[NextHop],

            nextHops =
                List.Distinct(
                    List.Transform(bufferedRouteRecords, (route as record) => route[NextHop]),
                    Comparer.Ordinal
                ),
            aggregateNextHop = (nextHop as text) as list =>
                let
                    group =
                        List.Select(
                            bufferedRouteRecords,
                            (route as record) => Comparer.Ordinal(route[NextHop], nextHop) = 0
                        ),
                    sortedGroup = List.Sort(group, compareRoutes)
                in
                    aggregateGroup(sortedGroup),
            aggregatedRoutes =
                List.Accumulate(
                    nextHops,
                    {},
                    (state as list, nextHop as text) => state & aggregateNextHop(nextHop)
                ),
            sortedAggregatedRoutes = List.Sort(aggregatedRoutes, compareRoutes),
            ascendingResult = List.Transform(sortedAggregatedRoutes, renderRoute),
            descendingMode = descending ?? false
        in
            if descendingMode then
                List.Reverse(ascendingResult)
            else
                ascendingResult,

    IpRouteAggregateArrayType =
        type function (
           routes as (type nullable list meta [
                Documentation.FieldCaption = "IPv4 route list",
                Documentation.FieldDescription = "A list of IPv4 route text values with optional next-hop suffixes; the list itself cannot be null, and null and empty items are ignored."
           ]),
           optional descending as (type nullable logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Return the aggregated route list in descending network order."
           ])
        ) as list meta [
            Documentation.Name = "IpRouteAggregateArray",
            Documentation.LongDescription = "Normalizes, sorts, deduplicates, removes contained routes, and joins aligned equal-size routes only within the same exact next-hop group. Returns a Power Query list of route text, preserves CIDR versus dotted-mask notation from the first route that survives each merge, rejects a null route list with a structured validation error, ignores null and empty items, rejects non-text or malformed routes, and treats null descending as ascending order.",
            Documentation.Examples = {
                [
                    Description = "Join adjacent routes with the same next hop.",
                    Code = "IpRouteAggregateArray({\"10.0.0.0/25 gw\", \"10.0.0.128/25 gw\"})",
                    Result = "{\"10.0.0.0/24 gw\"}"
                ],
                [
                    Description = "Keep adjacent routes with different next hops separate.",
                    Code = "IpRouteAggregateArray({\"10.0.0.0/25 gw-a\", \"10.0.0.128/25 gw-b\"})",
                    Result = "{\"10.0.0.0/25 gw-a\", \"10.0.0.128/25 gw-b\"}"
                ],
                [
                    Description = "Aggregate dotted-mask routes and reverse the result.",
                    Code = "IpRouteAggregateArray({\"192.0.2.0 255.255.255.128 next-hop\", \"192.0.2.128 255.255.255.128 next-hop\"}, true)",
                    Result = "{\"192.0.2.0 255.255.255.0 next-hop\"}"
                ]
            }
        ],

    IpRouteAggregateArrayDoc = Value.ReplaceType(IpRouteAggregateArray, IpRouteAggregateArrayType),

    /*
        Divides an IPv4 subnet into equal-sized child subnets.

        The VBA reference accepts an optional index and otherwise returns an
        Excel array. The JavaScript counterpart returns the complete list and
        preserves the input object's notation. M uses one stable list result;
        callers select a child with ordinary zero-based list indexing, such as
        IpDivideSubnetDoc("1.2.3.0/24", 2){1}. It follows the VBA baseline by
        normalizing host bits and rendering canonical CIDR text, regardless of
        whether the input used CIDR, dotted-mask, or unmasked notation.

        The optional descending parameter models the project's list export
        convention. Null and malformed subnet text, negative or fractional
        offsets, and a target prefix beyond /32 raise structured validation
        errors. Offset zero returns the normalized input network as a one-item
        list.

        No child-count ceiling is enforced. List member expressions are
        evaluated lazily, so indexing a single child (as in the second example
        below) stays cheap even for a large offset; materializing the entire
        list for a large offset is deliberately left to the caller, since that
        cost is inherent to requesting that many rows rather than a defect in
        this function. Descending order is generated directly (not via
        List.Reverse) so that laziness is preserved in both directions.

        Examples:
            IpDivideSubnetDoc("1.2.3.0/24", 2)
                = {"1.2.3.0/26", "1.2.3.64/26", "1.2.3.128/26", "1.2.3.192/26"}
            IpDivideSubnetDoc("1.2.3.64/24", 2){1}
                = "1.2.3.64/26"
            IpDivideSubnetDoc("1.2.3.0/24", 1, true)
                = {"1.2.3.128/25", "1.2.3.0/25"}
    */
    IpDivideSubnet = (
        subnet as nullable text,
        offset as nullable number,
        optional descending as nullable logical
    ) as list =>
        let
            fail = IpMakeFailure("IpDivideSubnet.InvalidInput"),

            parsedSubnet = IpSubnetParseWithAddressValue(subnet),
            prefixLength = parsedSubnet[PrefixLength],
            checkedOffset =
                if offset = null then
                    fail(
                        "Subnet division offset cannot be null.",
                        [Input = offset, Component = "offset", Expected = "whole number >= 0"]
                    )
                else if Number.RoundDown(offset) <> offset or offset < 0 then
                    fail(
                        "Subnet division offset must be a whole number greater than or equal to zero.",
                        [Input = offset, Component = "offset", Expected = "whole number >= 0"]
                    )
                else
                    offset,
            targetPrefixLength = prefixLength + checkedOffset,
            checkedTargetPrefixLength =
                if targetPrefixLength > 32 then
                    fail(
                        "Subnet division target prefix length must not exceed 32.",
                        [
                            Input = subnet,
                            Component = "prefix length",
                            PrefixLength = prefixLength,
                            Offset = checkedOffset,
                            TargetPrefixLength = targetPrefixLength,
                            Expected = "0..32"
                        ]
                    )
                else
                    targetPrefixLength,
            parentFacts = parsedSubnet[Facts],
            childFacts = IpSubnetFacts(parentFacts[NetworkKey], checkedTargetPrefixLength),
            networkValue = parentFacts[NetworkKey],
            childBlockSize = childFacts[BlockSize],
            childCount = Number.Power(2, checkedOffset),
            descendingMode = descending ?? false,
            childIndexes =
                if descendingMode then
                    List.Numbers(childCount - 1, childCount, -1)
                else
                    List.Numbers(0, childCount, 1)
        in
            List.Transform(
                childIndexes,
                (childIndex as number) as text =>
                    IpBinToStr(networkValue + childIndex * childBlockSize)
                        & "/"
                        & Number.ToText(checkedTargetPrefixLength, "0", "en-US")
            ),

    IpDivideSubnetType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 subnet in CIDR, dotted-mask, or unmasked notation. Host bits are cleared before division."
           ]),
           offset as (type nullable number meta [
                Documentation.FieldCaption = "Division offset",
                Documentation.FieldDescription = "A whole number of prefix bits to add; the target prefix must be at most 32."
           ]),
           optional descending as (type nullable logical meta [
                Documentation.FieldCaption = "Descending",
                Documentation.FieldDescription = "Reverse the generated child-subnet list; null means ascending order."
           ])
        ) as list meta [
            Documentation.Name = "IpDivideSubnet",
            Documentation.LongDescription = "Returns a Power Query list of canonical CIDR child subnets after clearing host bits from the input network. The offset adds prefix bits, so the list contains 2^offset children; callers can select one with zero-based list indexing. Null, malformed, non-integral, negative, and too-large offsets raise structured validation errors, and null descending means ascending output.",
            Documentation.Examples = {
                [
                    Description = "Divide a /24 into four /26 child networks.",
                    Code = "IpDivideSubnet(\"1.2.3.0/24\", 2)",
                    Result = "{\"1.2.3.0/26\", \"1.2.3.64/26\", \"1.2.3.128/26\", \"1.2.3.192/26\"}"
                ],
                [
                    Description = "Select the second child with normal zero-based list indexing.",
                    Code = "IpDivideSubnet(\"1.2.3.64/24\", 2){1}",
                    Result = "\"1.2.3.64/26\""
                ],
                [
                    Description = "Return two /25 children in descending order.",
                    Code = "IpDivideSubnet(\"1.2.3.0/24\", 1, true)",
                    Result = "{\"1.2.3.128/25\", \"1.2.3.0/25\"}"
                ]
            }
        ],

    IpDivideSubnetDoc = Value.ReplaceType(IpDivideSubnet, IpDivideSubnetType),

    /*
        Returns the number of IPv4 addresses in a subnet.

        The VBA and JavaScript references derive the size from the validated
        subnet prefix length as 2^(32 - prefix). M delegates parsing and
        validation to IpSubnetLen, then uses Number.Power for the bounded
        IPv4 arithmetic. An unmasked address is therefore a /32 subnet with
        one address; /0 returns the full 2^32 IPv4 address space.

        Null, malformed, invalid-prefix, and non-canonical-mask inputs raise
        the structured validation errors from IpSubnetParse through
        IpSubnetLen.

        Examples:
            IpSubnetSizeDoc("192.168.1.32/29") = 8
            IpSubnetSizeDoc("192.168.1.0 255.255.255.0") = 256
            IpSubnetSizeDoc("192.168.1.1") = 1
    */
    IpSubnetSize = (subnet as nullable text) as number =>
        let
            prefixLength = IpSubnetLen(subnet)
        in
            Number.Power(2, 32 - prefixLength),

    IpSubnetSizeType =
        type function (
           subnet as (type nullable text meta [
                Documentation.FieldCaption = "IPv4 subnet",
                Documentation.FieldDescription = "An IPv4 address with optional CIDR or dotted-mask notation."
           ])
        ) as number meta [
            Documentation.Name = "IpSubnetSize",
            Documentation.LongDescription = "Returns 2^(32 - prefix length) for a validated IPv4 subnet. CIDR, dotted-decimal mask, and unmasked address notation are supported; unmasked addresses are treated as /32. Null, malformed, invalid-prefix, and non-canonical-mask inputs raise structured validation errors.",
            Documentation.Examples = {
                [
                    Description = "Count addresses in a CIDR subnet.",
                    Code = "IpSubnetSize(\"192.168.1.32/29\")",
                    Result = "8"
                ],
                [
                    Description = "Count addresses in dotted-mask notation.",
                    Code = "IpSubnetSize(\"192.168.1.0 255.255.255.0\")",
                    Result = "256"
                ],
                [
                    Description = "Treat an unmasked address as a /32 subnet.",
                    Code = "IpSubnetSize(\"192.168.1.1\")",
                    Result = "1"
                ]
            }
        ],

    IpSubnetSizeDoc = Value.ReplaceType(IpSubnetSize, IpSubnetSizeType)
in
    [
        IpStrToBin = IpStrToBinDoc,
        IpBinToStr = IpBinToStrDoc,
        IpIsValid = IpIsValidDoc,
        IpIsPrivate = IpIsPrivateDoc,
        IpParse = IpParseDoc,
        IpParseRoute = IpParseRouteDoc,
        IpFindOverlappingSubnets = IpFindOverlappingSubnetsDoc,
        IpBuild = IpBuildDoc,
        IpAdd = IpAddDoc,
        IpAdd2 = IpAdd2Doc,
        IpComp = IpCompDoc,
        IpAnd = IpAndDoc,
        IpOr = IpOrDoc,
        IpXor = IpXorDoc,
        IpDiff = IpDiffDoc,
        IpGetByte = IpGetByteDoc,
        IpSetByte = IpSetByteDoc,
        IpSortArray = IpSortArrayDoc,
        IpInvertMask = IpInvertMaskDoc,
        IpWildMask = IpWildMaskDoc,
        IpRangeToCIDR = IpRangeToCIDRDoc,
        IpRouteAggregateArray = IpRouteAggregateArrayDoc,
        IpDivideSubnet = IpDivideSubnetDoc,
        IpSubnetAggregateArray = IpSubnetAggregateArrayDoc,
        IpSubnetSortArray = IpSubnetSortArrayDoc,
        IpSubtractSubnets = IpSubtractSubnetsDoc,
        IpCommonSubnets = IpCommonSubnetsDoc,
        IpSubnetVLookupAreas = IpSubnetVLookupAreasDoc,
        IpMaskLen = IpMaskLenDoc,
        IpMaskBin = IpMaskBinDoc,
        IpMask = IpMaskDoc,
        IpSubnetParse = IpSubnetParseDoc,
        IpSubnetToBin = IpSubnetToBinDoc,
        IpSubnetMatch = IpSubnetMatchDoc,
        IpSubnetJoin = IpSubnetJoinDoc,
        IpWithoutMask = IpWithoutMaskDoc,
        IpClearHostBits = IpClearHostBitsDoc,
        IpSubnetLen = IpSubnetLenDoc,
        IpSubnetSize = IpSubnetSizeDoc,
        IpIsInSubnet = IpIsInSubnetDoc,
        IpSubnetIsInSubnet = IpSubnetIsInSubnetDoc
    ]
