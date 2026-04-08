export const REGIONS = [
  { id: "scotland", label: "Scotland", number: 1, svgId: "_x31_._Scotland" },
  { id: "northern-ireland", label: "Northern Ireland", number: 2, svgId: "_x32_._Northern_Ireland" },
  { id: "wales", label: "Wales", number: 3, svgId: "_x33_._Wales" },
  { id: "north-east", label: "North East", number: 4, svgId: "_x34_._North_East" },
  { id: "north-west", label: "North West", number: 5, svgId: "_x35_._North_West" },
  { id: "yorkshire-humber", label: "Yorkshire and the Humber", number: 6, svgId: "_x36_._Yorkshire_and_the_Humber" },
  { id: "west-midlands", label: "West Midlands", number: 7, svgId: "_x37_._West_Midlands" },
  { id: "east-midlands", label: "East Midlands", number: 8, svgId: "_x38_._East_Midlands" },
  { id: "south-west", label: "South West", number: 9, svgId: "_x39_._South_West" },
  { id: "south-east", label: "South East", number: 10, svgId: "_x31_0._South_East" },
  { id: "east-of-england", label: "East of England", number: 11, svgId: "_x31_1._East_of_England" },
  { id: "greater-london", label: "Greater London", number: 12, svgId: "_x31_2._Greater_London" },
  { id: "ireland", label: "Ireland", number: 13, svgId: "_x31_3._Ireland" }
];

export const FILTER_OPTIONS = [
  { id: "national-company", label: "National Company" },
  { id: "pasma", label: "PASMA" },
  { id: "ipaf", label: "IPAF" },
  { id: "cscs", label: "CSCS" },
  { id: "vetted", label: "Vetted" }
];

export const REGION_MAP = Object.fromEntries(REGIONS.map((region) => [region.id, region]));

export const DEFAULT_INSTALLER_FORM = {
  id: "",
  name: "",
  company: "",
  phone: "",
  email: "",
  address: "",
  notes: "",
  rating: 0,
  regions: [],
  tags: []
};
