import { REGION_MAP, REGIONS } from "./directoryData";

const POSTCODE_REGION_RULES = [
  { prefixes: ["AB", "DD", "DG", "EH", "FK", "G", "HS", "IV", "KA27", "KA28", "KW", "KY", "ML", "PA", "PH", "TD"], regionId: "scotland" },
  { prefixes: ["BT"], regionId: "northern-ireland" },
  { prefixes: ["CF", "LD", "LL", "NP", "SA", "SY16", "SY17", "SY18", "SY19", "SY20", "SY21", "SY22", "SY23", "SY24", "SY25"], regionId: "wales" },
  { prefixes: ["DH", "DL", "NE", "SR", "TS"], regionId: "north-east" },
  { prefixes: ["BB", "BL", "CA", "CH", "CW", "FY", "L", "LA", "M", "OL", "PR", "SK", "WA", "WN"], regionId: "north-west" },
  { prefixes: ["BD", "DN", "HD", "HG", "HU", "HX", "LS", "S", "WF", "YO"], regionId: "yorkshire-humber" },
  { prefixes: ["B", "CV", "DY", "HR", "ST", "SY1", "SY2", "SY3", "SY4", "SY5", "SY6", "SY7", "SY8", "SY9", "SY10", "SY11", "SY12", "SY13", "SY14", "SY15", "TF", "WR", "WS", "WV"], regionId: "west-midlands" },
  { prefixes: ["DE", "LE", "LN", "NG", "NN", "PE"], regionId: "east-midlands" },
  { prefixes: ["BA", "BS", "EX", "GL", "PL", "SN", "TA", "TQ", "TR"], regionId: "south-west" },
  { prefixes: ["BN", "CT", "GU", "HP", "ME", "MK", "OX", "PO", "RG", "RH", "SL", "SO", "TN"], regionId: "south-east" },
  { prefixes: ["AL", "CB", "CM", "CO", "IP", "LU", "NR", "SG", "SS"], regionId: "east-of-england" },
  { prefixes: ["E", "EC", "N", "NW", "SE", "SW", "W", "WC"], regionId: "greater-london" },
  { prefixes: ["D"], regionId: "ireland" }
];

const PLACE_REGION_RULES = [
  { keywords: ["SOUTHAMPTON", "PORTSMOUTH", "BRIGHTON", "READING", "OXFORD", "MILTON KEYNES", "MAIDSTONE", "CANTERBURY", "GUILDFORD"], regionId: "south-east" },
  { keywords: ["BRISTOL", "EXETER", "PLYMOUTH", "TRURO", "TAUNTON", "BATH", "GLOUCESTER", "CHELTENHAM", "BOURNEMOUTH", "POOLE", "WEYMOUTH", "SWINDON"], regionId: "south-west" },
  { keywords: ["NORWICH", "IPSWICH", "COLCHESTER", "CAMBRIDGE", "LUTON", "SOUTHEND", "CHELMSFORD", "PETERBOROUGH"], regionId: "east-of-england" },
  { keywords: ["LONDON"], regionId: "greater-london" },
  { keywords: ["BIRMINGHAM", "COVENTRY", "WOLVERHAMPTON", "DUDLEY", "HEREFORD", "TELFORD", "WORCESTER", "WALSALL"], regionId: "west-midlands" },
  { keywords: ["NOTTINGHAM", "LEICESTER", "DERBY", "LINCOLN", "NORTHAMPTON"], regionId: "east-midlands" },
  { keywords: ["MANCHESTER", "LIVERPOOL", "PRESTON", "BLACKPOOL", "LANCASTER", "BOLTON", "WARRINGTON", "CARLISLE", "CHESTER"], regionId: "north-west" },
  { keywords: ["LEEDS", "SHEFFIELD", "YORK", "HULL", "HUDDERSFIELD", "BRADFORD", "WAKEFIELD", "DONCASTER", "ROTHERHAM"], regionId: "yorkshire-humber" },
  { keywords: ["NEWCASTLE", "SUNDERLAND", "MIDDLESBROUGH", "DURHAM", "DARLINGTON", "HARTLEPOOL"], regionId: "north-east" },
  { keywords: ["CARDIFF", "SWANSEA", "NEWPORT", "WREXHAM", "ABERYSTWYTH", "LLANDUDNO"], regionId: "wales" },
  { keywords: ["GLASGOW", "EDINBURGH", "ABERDEEN", "DUNDEE", "INVERNESS", "PERTH", "STIRLING"], regionId: "scotland" },
  { keywords: ["BELFAST", "LISBURN", "NEWRY", "DERRY"], regionId: "northern-ireland" },
  { keywords: ["DUBLIN", "CORK", "GALWAY", "LIMERICK", "WATERFORD"], regionId: "ireland" }
];

export function createId() {
  return globalThis.crypto?.randomUUID?.() ?? `installer-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

export function normalizeInstaller(installer) {
  return {
    id: installer.id || createId(),
    name: installer.name || "",
    company: installer.company || "",
    phone: installer.phone || "",
    email: installer.email || "",
    address: installer.address || "",
    notes: installer.notes || "",
    rating: Number.isFinite(Number(installer.rating)) ? Number(installer.rating) : 0,
    regions: Array.isArray(installer.regions) ? installer.regions : [],
    tags: Array.isArray(installer.tags) ? installer.tags : []
  };
}

export function getRegionStrength(count, maxCount) {
  if (!count || !maxCount) return "#d8d0dd";
  const ratio = count / maxCount;
  if (ratio >= 0.85) return "#5f3c74";
  if (ratio >= 0.65) return "#75508c";
  if (ratio >= 0.45) return "#8a69a1";
  if (ratio >= 0.25) return "#b39ac1";
  return "#d9cde1";
}

export function getRegionFromAddress(input) {
  const raw = String(input || "").trim();
  if (!raw) return null;
  const text = raw.toUpperCase();
  const regionMatch = REGIONS.find((region) => text.includes(region.label.toUpperCase()));
  if (regionMatch) return regionMatch;

  for (const rule of PLACE_REGION_RULES) {
    if (rule.keywords.some((keyword) => text.includes(keyword))) {
      return REGION_MAP[rule.regionId];
    }
  }

  const compact = text.replace(/\s+/g, "");
  const postcodeMatch = compact.match(/\b[A-Z]{1,2}\d[A-Z\d]?\d[A-Z]{2}\b|\b[A-Z]{1,2}\d[A-Z\d]?\b/);
  const candidate = postcodeMatch ? postcodeMatch[0] : null;

  if (candidate && /\d/.test(candidate)) {
    for (const rule of POSTCODE_REGION_RULES) {
      if (rule.prefixes.some((prefix) => candidate.startsWith(prefix))) {
        return REGION_MAP[rule.regionId];
      }
    }
  }

  return null;
}
