using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace readexcelx
{
    public  class countryDicValue
    {
        public Dictionary<string, string> countryDic = new Dictionary<string, string>();

        public  Dictionary<string, string> getCountries()
        {
            return countryDic;
        }

        public void setCountries()
        {
           
            try
            {
                countryDic.Add("1", "Afghanistan (1)");
                countryDic.Add("274", "Aland Islands (274)");
                countryDic.Add("2", "Albania (2)");
                countryDic.Add("3", "Algeria (3)");
                countryDic.Add("4", "American Samoa (4)");
                countryDic.Add("5", "Andorra (5)");
                countryDic.Add("6", "Angola (6)");
                countryDic.Add("7", "Anguilla (7)");
                countryDic.Add("8", "Antarctica (8)");
                countryDic.Add("9", "Antigua and Barbuda (9)");
                countryDic.Add("10", "Argentina (10)");
                countryDic.Add("11", "Armenia (11)");
                countryDic.Add("12", "Aruba (12)");
                countryDic.Add("13", "Ashmore and Cartier Islands (13)");
                countryDic.Add("14", "Australia (14)");
                countryDic.Add("15", "Austria (15)");
                countryDic.Add("16", "Azerbaijan (16)");
                countryDic.Add("17", "Bahamas (17)");
                countryDic.Add("18", "Bahrain (18)");
                countryDic.Add("19", "Baker Island (19)");
                countryDic.Add("20", "Bangladesh (20)");
                countryDic.Add("21", "Barbados (21)");
                countryDic.Add("22", "Bassas Da India (22)");
                countryDic.Add("23", "Belarus (23)");
                countryDic.Add("24", "Belgium (24)");
                countryDic.Add("25", "Belize (25)");
                countryDic.Add("26", "Benin (26)");
                countryDic.Add("27", "Bermuda (27)");
                countryDic.Add("28", "Bhutan (28)");
                countryDic.Add("29", "BO (29)");
                countryDic.Add("30", "Bolivia (30)");
                countryDic.Add("275", "Bonaire, St. Eustatius and Saba (275)");
                countryDic.Add("31", "Bosnia and Herzegovina (31)");
                countryDic.Add("32", "Botswana (32)");
                countryDic.Add("33", "Bouvet Island (33)");
                countryDic.Add("34", "Brazil (34)");
                countryDic.Add("35", "British Indian Ocean Territory (35)");
                countryDic.Add("37", "Brunei Darussalam (37)");
                countryDic.Add("38", "Bulgaria (38)");
                countryDic.Add("39", "Burkina Faso (39)");
                countryDic.Add("41", "Burundi (41)");
                countryDic.Add("42", "Cambodia (42)");
                countryDic.Add("43", "Cameroon (43)");
                countryDic.Add("44", "Canada (44)");
                countryDic.Add("45", "Cape Verde (45)");
                countryDic.Add("46", "Cayman Islands (46)");
                countryDic.Add("47", "Central African Republic (47)");
                countryDic.Add("48", "Chad (48)");
                countryDic.Add("49", "Chile (49)");
                countryDic.Add("50", "China (50)");
                countryDic.Add("51", "Christmas Island (51)");
                countryDic.Add("52", "Clipperton Island (52)");
                countryDic.Add("53", "Cocos (Keeling) Islands (53)");
                countryDic.Add("54", "Colombia (54)");
                countryDic.Add("55", "Comoros (55)");
                countryDic.Add("57", "Congo, Democratic Republic of the (57)");
                countryDic.Add("56", "Congo, Republic of the (56)");
                countryDic.Add("58", "Cook Islands (58)");
                countryDic.Add("59", "Coral Sea Islands (59)");
                countryDic.Add("60", "Costa Rica (60)");
                countryDic.Add("61", "Côte D'ivoire (61)");
                countryDic.Add("62", "Croatia (62)");
                countryDic.Add("63", "Cuba (63)");
                countryDic.Add("276", "Curacao (276)");
                countryDic.Add("64", "Cyprus (64)");
                countryDic.Add("65", "Czech Republic (65)");
                countryDic.Add("66", "Denmark (66)");
                countryDic.Add("67", "Djibouti (67)");
                countryDic.Add("68", "Dominica (68)");
                countryDic.Add("69", "Dominican Republic (69)");
                countryDic.Add("70", "Dubai (70)");
                countryDic.Add("272", "East Timor (272)");
                countryDic.Add("71", "Ecuador (71)");
                countryDic.Add("72", "Egypt (72)");
                countryDic.Add("73", "El Salvador (73)");
                countryDic.Add("284", "England (284)");
                countryDic.Add("74", "Equatorial Guinea (74)");
                countryDic.Add("75", "Eritrea (75)");
                countryDic.Add("76", "Estonia (76)");
                countryDic.Add("77", "Ethiopia (77)");
                countryDic.Add("78", "Europa Island (78)");
                countryDic.Add("79", "Falkland Islands (Malvinas) (79)");
                countryDic.Add("80", "Faroe Islands (80)");
                countryDic.Add("81", "Fiji (81)");
                countryDic.Add("82", "Finland (82)");
                countryDic.Add("83", "France (83)");
                countryDic.Add("84", "French Guiana (84)");
                countryDic.Add("85", "French Polynesia (85)");
                countryDic.Add("86", "French Southern Territories (86)");
                countryDic.Add("87", "Gabon (87)");
                countryDic.Add("89", "Gaza Strip (89)");
                countryDic.Add("90", "Georgia (90)");
                countryDic.Add("91", "Germany (91)");
                countryDic.Add("92", "Ghana (92)");
                countryDic.Add("93", "Gibraltar (93)");
                countryDic.Add("94", "Glorioso Islands (94)");
                countryDic.Add("95", "Greece (95)");
                countryDic.Add("96", "Greenland (96)");
                countryDic.Add("97", "Grenada (97)");
                countryDic.Add("98", "Guadeloupe (98)");
                countryDic.Add("99", "Guam (99)");
                countryDic.Add("100", "Guatemala (100)");
                countryDic.Add("101", "Guernsey (101)");
                countryDic.Add("102", "Guinea (102)");
                countryDic.Add("103", "Guinea-Bissau (103)");
                countryDic.Add("104", "Guyana (104)");
                countryDic.Add("105", "Haiti (105)");
                countryDic.Add("106", "Heard Island and Mcdonald Islands (106)");
                countryDic.Add("258", "Holy See (Vatican City State) (258)");
                countryDic.Add("107", "Honduras (107)");
                countryDic.Add("108", "Hong Kong (108)");
                countryDic.Add("109", "Howland Island (109)");
                countryDic.Add("110", "Hungary (110)");
                countryDic.Add("111", "Iceland (111)");
                countryDic.Add("112", "India (112)");
                countryDic.Add("113", "Indonesia (113)");
                countryDic.Add("114", "International (114)");
                countryDic.Add("115", "Iran, Islamic Republic Of (115)");
                countryDic.Add("116", "Iraq (116)");
                countryDic.Add("117", "Ireland (117)");
                countryDic.Add("118", "Isle Of Man (118)");
                countryDic.Add("119", "Israel (119)");
                countryDic.Add("120", "Italy (120)");
                countryDic.Add("121", "Ivory Coast (121)");
                countryDic.Add("122", "Jamaica (122)");
                countryDic.Add("123", "Jan Mayen (123)");
                countryDic.Add("124", "Japan (124)");
                countryDic.Add("125", "Jarvis Island (125)");
                countryDic.Add("126", "Jersey (126)");
                countryDic.Add("127", "Johnston Atoll (127)");
                countryDic.Add("128", "Jordan (128)");
                countryDic.Add("129", "Juan De Nova Island (129)");
                countryDic.Add("130", "Kazakhstan (130)");
                countryDic.Add("131", "Kenya (131)");
                countryDic.Add("132", "Kingman Reef (132)");
                countryDic.Add("133", "Kiribati (133)");
                countryDic.Add("134", "Kuwait (134)");
                countryDic.Add("135", "Kyrgyzstan (135)");
                countryDic.Add("136", "Laos (136)");
                countryDic.Add("137", "Latvia (137)");
                countryDic.Add("138", "Lebanon (138)");
                countryDic.Add("139", "Lesotho (139)");
                countryDic.Add("140", "Liberia (140)");
                countryDic.Add("141", "Libya (141)");
                countryDic.Add("142", "Liechtenstein (142)");
                countryDic.Add("143", "Lithuania (143)");
                countryDic.Add("144", "Luxembourg (144)");
                countryDic.Add("145", "Macau (145)");
                countryDic.Add("146", "Macedonia (146)");
                countryDic.Add("147", "Macedonia, The Former Yugoslav Republic Of (147)");
                countryDic.Add("148", "Madagascar (148)");
                countryDic.Add("149", "Malawi (149)");
                countryDic.Add("150", "Malaysia (150)");
                countryDic.Add("151", "Maldives (151)");
                countryDic.Add("152", "Mali (152)");
                countryDic.Add("153", "Malta (153)");
                countryDic.Add("154", "Marshall Islands (154)");
                countryDic.Add("155", "Martinique (155)");
                countryDic.Add("156", "Mauritania (156)");
                countryDic.Add("157", "Mauritius (157)");
                countryDic.Add("158", "Mayotte (158)");
                countryDic.Add("159", "Mexico (159)");
                countryDic.Add("160", "Micronesia, Federated States Of (160)");
                countryDic.Add("161", "Midway Islands (161)");
                countryDic.Add("162", "Moldova, Republic Of (162)");
                countryDic.Add("163", "Monaco (163)");
                countryDic.Add("164", "Mongolia (164)");
                countryDic.Add("165", "Montenegro (165)");
                countryDic.Add("166", "Montserrat (166)");
                countryDic.Add("167", "Morocco (167)");
                countryDic.Add("168", "Mozambique (168)");
                countryDic.Add("40", "Myanmar (40)");
                countryDic.Add("169", "Myanmar (Burma) (169)");
                countryDic.Add("170", "Namibia (170)");
                countryDic.Add("171", "Nauru (171)");
                countryDic.Add("173", "Nepal (173)");
                countryDic.Add("174", "Netherlands (174)");
                countryDic.Add("175", "Netherlands Antilles (175)");
                countryDic.Add("176", "New Caledonia (176)");
                countryDic.Add("177", "New Zealand (177)");
                countryDic.Add("178", "Nicaragua (178)");
                countryDic.Add("179", "Niger (179)");
                countryDic.Add("180", "Nigeria (180)");
                countryDic.Add("181", "Niue (181)");
                countryDic.Add("182", "Norfolk Island (182)");
                countryDic.Add("183", "North Korea (183)");
                countryDic.Add("184", "Northern Mariana Islands (184)");
                countryDic.Add("185", "Norway (185)");
                countryDic.Add("186", "Oman (186)");
                countryDic.Add("187", "Pakistan (187)");
                countryDic.Add("188", "Palau (188)");
                countryDic.Add("264", "Palestinian Territory, Occupied (264)");
                countryDic.Add("271", "Palestinian Territory, Occupied (271)");
                countryDic.Add("189", "Palmyra Atoll (189)");
                countryDic.Add("190", "Panama (190)");
                countryDic.Add("191", "Papua New Guinea (191)");
                countryDic.Add("192", "Paracel Islands (192)");
                countryDic.Add("193", "Paraguay (193)");
                countryDic.Add("194", "Peru (194)");
                countryDic.Add("195", "Philippines (195)");
                countryDic.Add("196", "Pitcairn (196)");
                countryDic.Add("197", "Poland (197)");
                countryDic.Add("198", "Portugal (198)");
                countryDic.Add("199", "Puerto Rico (199)");
                countryDic.Add("200", "Qatar (200)");
                countryDic.Add("283", "Republic of Kosovo (283)");
                countryDic.Add("210", "Republic of Serbia (210)");
                countryDic.Add("88", "Republic of the Gambia (88)");
                countryDic.Add("201", "Reunion (201)");
                countryDic.Add("202", "Romania (202)");
                countryDic.Add("203", "Russia (203)");
                countryDic.Add("204", "Rwanda (204)");
                countryDic.Add("224", "Saint Helena (224)");
                countryDic.Add("225", "Saint Kitts and Nevis (225)");
                countryDic.Add("226", "Saint Lucia (226)");
                countryDic.Add("277", "Saint Maarten   (277)");
                countryDic.Add("227", "Saint Pierre and Miquelon (227)");
                countryDic.Add("205", "Samoa (205)");
                countryDic.Add("206", "San Marino (206)");
                countryDic.Add("207", "Sao Tome and Principe (207)");
                countryDic.Add("208", "Saudi Arabia (208)");
                countryDic.Add("285", "Scotland (285)");
                countryDic.Add("209", "Senegal (209)");
                countryDic.Add("211", "Seychelles (211)");
                countryDic.Add("282", "Sicily (282)");
                countryDic.Add("212", "Sierra Leone (212)");
                countryDic.Add("213", "Singapore (213)");
                countryDic.Add("214", "Slovakia (214)");
                countryDic.Add("215", "Slovenia (215)");
                countryDic.Add("216", "Solomon Islands (216)");
                countryDic.Add("217", "Somalia (217)");
                countryDic.Add("218", "South Africa (218)");
                countryDic.Add("219", "South Georgia and The South Sandwich Islands (219)");
                countryDic.Add("220", "South Korea (220)");
                countryDic.Add("278", "South Sudan   (278)");
                countryDic.Add("221", "Spain (221)");
                countryDic.Add("222", "Spratly Islands (222)");
                countryDic.Add("223", "Sri Lanka (223)");
                countryDic.Add("279", "St. Barthelemy        (279)");
                countryDic.Add("280", "St. Martin (280)");
                countryDic.Add("228", "St. Vincent and The Grenadines (228)");
                countryDic.Add("229", "Sudan (229)");
                countryDic.Add("230", "Suriname (230)");
                countryDic.Add("231", "Svalbard and Jan Mayen (231)");
                countryDic.Add("232", "Swaziland (232)");
                countryDic.Add("233", "Sweden (233)");
                countryDic.Add("234", "Switzerland (234)");
                countryDic.Add("235", "Syrian Arab Republic (235)");
                countryDic.Add("236", "Taiwan (236)");
                countryDic.Add("237", "Tajikistan (237)");
                countryDic.Add("238", "Tanzania (238)");
                countryDic.Add("239", "Thailand (239)");
                countryDic.Add("281", "Timor-Leste   (281)");
                countryDic.Add("240", "Togo (240)");
                countryDic.Add("241", "Tokelau (241)");
                countryDic.Add("242", "Tonga (242)");
                countryDic.Add("243", "Trinidad and Tobago (243)");
                countryDic.Add("244", "Tromelin Island (244)");
                countryDic.Add("245", "Tunisia (245)");
                countryDic.Add("246", "Turkey (246)");
                countryDic.Add("247", "Turkmenistan (247)");
                countryDic.Add("248", "Turks and Caicos Islands (248)");
                countryDic.Add("249", "Tuvalu (249)");
                countryDic.Add("250", "Uganda (250)");
                countryDic.Add("251", "Ukraine (251)");
                countryDic.Add("252", "United Arab Emirates (252)");
                countryDic.Add("253", "United Kingdom (253)");
                countryDic.Add("273", "United States Minor Outlying Islands (273)");
                countryDic.Add("172", "United States Minor Outlying Islands (172)");
                countryDic.Add("254", "United States Of America (254)");
                countryDic.Add("255", "Uruguay (255)");
                countryDic.Add("256", "Uzbekistan (256)");
                countryDic.Add("257", "Vanuatu (257)");
                countryDic.Add("259", "Venezuela (259)");
                countryDic.Add("260", "Vietnam (260)");
                countryDic.Add("36", "Virgin Islands, British (36)");
                countryDic.Add("261", "Virgin Islands, U.S. (261)");
                countryDic.Add("262", "Wake Island (262)");
                countryDic.Add("263", "Wallis and Futuna (263)");
                countryDic.Add("265", "Western Sahara (265)");
                countryDic.Add("266", "Western Samoa (266)");
                countryDic.Add("267", "Yemen (267)");
                countryDic.Add("268", "Yugoslavia (268)");
                countryDic.Add("269", "Zambia (269)");
                countryDic.Add("270", "Zimbabwe (270)");



            }
            catch (Exception ex)
            {
               
                throw;
            }
          
        }
    }

    public class DefaultJsonSerializer : JsonSerializerSettings
    {
        public DefaultJsonSerializer()
        {
            NullValueHandling = NullValueHandling.Ignore;
            DefaultValueHandling = DefaultValueHandling.Ignore;

        }
    }

    class BBDataExtractClass
    {
        public class BenefitsRules
        {
            public Int32? Points { get; set; }
            public string Attribute { get; set; }
            public bool IsRecommended { get; set; }
            public string BenefitID { get; set; }
            public string ServiceType { get; set; }
        }

        public class Rules
        {
            public List<int> JobCode { get; set; }
            public bool IntraRegionalFlg { get; set; }
            public int TotalPoints { get; set; }
            public int RuleID { get; set; }
            public List<string> DestinationCountryCd { get; set; }
            public List<string> DepartureCountryCd { get; set; }
            public string SpouseMoving { get; set; }
            public string FamilyMoving { get; set; }
            public List<BenefitsRules> Benefits { get; set; }
        }

        public class policyRules
        {
            public int PolicyID { get; set; }
            public string PolicyName { get; set; }
            public List<Rules> Rules { get; set; }
        }
        public class ClientRules
        {
            public string _id { get; set; }

            public string ClientNo { get; set; }
            public List<policyRules> Policies { get; set; }
        }

        public class Client
        {
            public string _id { get; set; }
            public Config Config { get; set; }
            public string ClientNo { get; set; }
            public List<Policies> Policies { get; set; }
        }

        public class Config
        {
            public bool DisableBeforeWelcomeCall { get; set; }
            public List<string> CategorySortOrder { get; set; }
            public List<string> BenefitOrder { get; set; }
        }
        public class CashOut
        {
            public List<string> BasedOn { get; set; }
            public List<CashOutRules> CashOutRules { get; set; }
        }
        public class CashOutRules
        {
            public List<double> PointsCashedOut { get; set; }
            public List<int> JobCode { get; set; }
            public double Amount { get; set; }
            public string Currency { get; set; }
        }
        public class Policies
        {
            public string PolicyID { get; set; }
            public string PolicyName { get; set; }
            public string PDFSalutation { get; set; }
            public List<BenefitCards> Benefits { get; set; }
            public CashOut CashOut { get; set; }
        }

        public class BenefitCards 
        {
            public string Attribute { get; set; }

            public string Category { get; set; }
            public string ServiceType { get; set; }
            public bool IsRecommended { get; set; }
            public string Points { get; set; }
            public List<string> OrBenefits { get; set; }
            public List<string> AndBenefits { get; set; }
            public string LastUpdatedBy { get; set; }
            public bool HasQuantity { get; set; }
            public double CashOutValue { get; set; }
            public string BenefitID { get; set; }
            public string ClientBenefitDesc { get; set; }
            public string ClientBenefitTitle { get; set; }
            public string ClientNo { get; set; }
            public string ImageURL { get; set; }
            public string ProductNo { get; set; }
            public string SubProductNo { get; set; }
            public string ProdName { get; set; }
            public string SubProdName { get; set; }
            public double cardSequence { get; set; }
            public bool ConsultantOnly { get; set; }
            public bool Hide { get; set; }

        }


    }
}
