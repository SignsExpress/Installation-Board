import {Fragment,useEffect,useMemo,useRef,useState} from "react";
import {createPortal} from "react-dom";
const canEditVanEstimator=user=>user.role!=="member";
function MainNavBar(){return null;}
const VAN_REFERENCE_TYRE_DIAMETER_MM = 686.5;

const VAN_REFERENCE_TYRE_DIAMETER_UNITS = 194.62;

const VAN_ESTIMATOR_TEMPLATE = {
  id: "medium-van-transit-custom",
  sizeName: "Medium Van",
  exampleName: "Ford Transit Custom",
  src: "/vans/ford-transit-custom-swb.svg",
  scaleFactor: VAN_REFERENCE_TYRE_DIAMETER_MM / VAN_REFERENCE_TYRE_DIAMETER_UNITS,
  viewBox: { x: 0, y: 0, width: 2280.56, height: 1298.24 },
  pricingDefaultsVersion: 1
};

const VEHICLE_TEMPLATE_OPTIONS = [
  VAN_ESTIMATOR_TEMPLATE,
  {
    id: "large-van-ford-transit-panel-lwb",
    sizeName: "Large Van",
    exampleName: "Ford Transit Panel Van LWB",
    src: "/vans/ford-transit-panel-lwb.svg",
    scaleFactor: 743 / 210.56,
    scaleReferenceLayer: "742mm_Dia",
    tyreReferenceSelector: "#Artwork path.st6",
    tyreReferenceDiameterMm: 743,
    artworkScale: 0.1,
    viewBox: { x: 0, y: 0, width: 2595.02, height: 1624.19 },
    pricingDefaultsVersion: 2,
    pricingSettings: {
      standardVinylRate: 86,
      wrapRateStart: 110,
      wrapRateFloor: 80,
      wrapRateTaper: 32,
      labourSellRate: 160,
      standardSmallHoursPerM2: 0.7,
      standardSmallMinHours: 1,
      standardSmallMaxHours: 1.7,
      standardLargeHoursPerM2: 0.52,
      wrapLabourStartHoursPerM2: 1.02,
      wrapLabourFloorHoursPerM2: 0.66,
      wrapLabourTaper: 0.42,
      sectionFactors: {
        one: 0.92,
        twoToThree: 1.03,
        fourToFive: 1.12,
        moreThanFive: 1.2
      },
      difficultyFactors: {
        flat: 1,
        light_curve: 1.1,
        normal_wrap_curve: 1.2,
        deep_recess: 1.32
      },
      marketAnchors: {
        c0: 0,
        c05: 300,
        c10: 500,
        c15: 800,
        c22: 1300,
        c35: 1900,
        c55: 2850,
        c85: 3850,
        c100: 4100
      },
      materialMultipliers: {
        standard: 1,
        contra: 1.2,
        reflective: 2
      },
      blendWeights: {
        noWrap: { calculated: 0.46, anchor: 0.54 },
        wrapUnder35: { calculated: 0.42, anchor: 0.58 },
        wrapUnder70: { calculated: 0.28, anchor: 0.72 },
        wrapFull: { calculated: 0.15, anchor: 0.85 }
      },
      minPrice: 300,
      minAnyWrapPrice: 700,
      minPartialWrapPrice: 1050,
      minFullWrapPrice: 2500
    }
  },
  {
    id: "luton-van-vauxhall-movano",
    sizeName: "Luton Van",
    exampleName: "Vauxhall Movano Luton",
    src: "/vans/vauxhall-movano-luton.svg",
    scaleFactor: 740.98 / 210.06,
    tyreReferenceSelector: "g[id='_x37_40.98mm'] path",
    tyreReferenceDiameterMm: 740.98,
    artworkScale: 0.1,
    viewBox: { x: 0, y: -120, width: 2668.62, height: 2113.41 },
    pricingDefaultsVersion: 1
  }
];

const VEHICLE_GRAPHICS_PRICING = {
  standardVinylRate: 84,
  wrapRateStart: 107,
  wrapRateFloor: 77,
  wrapRateTaper: 34,
  labourSellRate: 160,
  standardSmallHoursPerM2: 0.68,
  standardSmallMinHours: 1,
  standardSmallMaxHours: 1.6,
  standardLargeHoursPerM2: 0.5,
  wrapLabourStartHoursPerM2: 0.98,
  wrapLabourFloorHoursPerM2: 0.62,
  wrapLabourTaper: 0.44,
  sectionFactors: {
    one: 0.92,
    twoToThree: 1.02,
    fourToFive: 1.1,
    moreThanFive: 1.18
  },
  difficultyFactors: {
    flat: 1,
    light_curve: 1.09,
    normal_wrap_curve: 1.18,
    deep_recess: 1.3
  },
  marketAnchors: {
    c0: 0,
    c05: 275,
    c10: 450,
    c15: 725,
    c22: 1175,
    c35: 1750,
    c55: 2600,
    c85: 3500,
    c100: 3750
  },
  blendWeights: {
    noWrap: { calculated: 0.48, anchor: 0.52 },
    wrapUnder35: { calculated: 0.44, anchor: 0.56 },
    wrapUnder70: { calculated: 0.32, anchor: 0.68 },
    wrapFull: { calculated: 0.18, anchor: 0.82 }
  },
  minPrice: 275,
  minAnyWrapPrice: 675,
  minPartialWrapPrice: 975,
  minFullWrapPrice: 2250,
  materialMultipliers: {
    standard: 1,
    contra: 1.2,
    reflective: 2
  }
};

const VEHICLE_PRICING_STORAGE_KEY = "standalone-pricing-" + window.__COMPANY_ID;

const VEHICLE_PRICE_TRAINING_BANK_STORAGE_KEY = "standalone-training-" + window.__COMPANY_ID;

const SMART_PRICE_IMPORT_FIELDS = {
  "standard vinyl rate": ["standardVinylRate"],
  "contra multiplier": ["materialMultipliers", "contra"],
  "reflective multiplier": ["materialMultipliers", "reflective"],
  "wrap start rate": ["wrapRateStart"],
  "wrap floor rate": ["wrapRateFloor"],
  "wrap material taper": ["wrapRateTaper"],
  "small std hrs/m2": ["standardSmallHoursPerM2"],
  "small std hrs m2": ["standardSmallHoursPerM2"],
  "small standard hrs/m2": ["standardSmallHoursPerM2"],
  "small std min hrs": ["standardSmallMinHours"],
  "small std max hrs": ["standardSmallMaxHours"],
  "large std hrs/m2": ["standardLargeHoursPerM2"],
  "large std hrs m2": ["standardLargeHoursPerM2"],
  "large standard hrs/m2": ["standardLargeHoursPerM2"],
  "wrap start hrs/m2": ["wrapLabourStartHoursPerM2"],
  "wrap start hrs m2": ["wrapLabourStartHoursPerM2"],
  "wrap floor hrs/m2": ["wrapLabourFloorHoursPerM2"],
  "wrap floor hrs m2": ["wrapLabourFloorHoursPerM2"],
  "wrap labour taper": ["wrapLabourTaper"],
  "1 section factor": ["sectionFactors", "one"],
  "one section factor": ["sectionFactors", "one"],
  "2-3 sections factor": ["sectionFactors", "twoToThree"],
  "2 3 sections factor": ["sectionFactors", "twoToThree"],
  "4-5 sections factor": ["sectionFactors", "fourToFive"],
  "4 5 sections factor": ["sectionFactors", "fourToFive"],
  "6+ sections factor": ["sectionFactors", "moreThanFive"],
  "6 sections factor": ["sectionFactors", "moreThanFive"],
  "normal wrap difficulty": ["difficultyFactors", "normal_wrap_curve"],
  "labour sell/hr": ["labourSellRate"],
  "labour sell hr": ["labourSellRate"],
  "labor sell/hr": ["labourSellRate"],
  "anchor 0%": ["marketAnchors", "c0"],
  "anchor 0": ["marketAnchors", "c0"],
  "anchor 5%": ["marketAnchors", "c05"],
  "anchor 5": ["marketAnchors", "c05"],
  "anchor 10%": ["marketAnchors", "c10"],
  "anchor 10": ["marketAnchors", "c10"],
  "anchor 15%": ["marketAnchors", "c15"],
  "anchor 15": ["marketAnchors", "c15"],
  "anchor 22%": ["marketAnchors", "c22"],
  "anchor 22": ["marketAnchors", "c22"],
  "anchor 35%": ["marketAnchors", "c35"],
  "anchor 35": ["marketAnchors", "c35"],
  "anchor 55%": ["marketAnchors", "c55"],
  "anchor 55": ["marketAnchors", "c55"],
  "anchor 85%": ["marketAnchors", "c85"],
  "anchor 85": ["marketAnchors", "c85"],
  "anchor 100%": ["marketAnchors", "c100"],
  "anchor 100": ["marketAnchors", "c100"],
  "no-wrap calc weight": ["blendWeights", "noWrap", "calculated"],
  "no wrap calc weight": ["blendWeights", "noWrap", "calculated"],
  "no-wrap anchor weight": ["blendWeights", "noWrap", "anchor"],
  "no wrap anchor weight": ["blendWeights", "noWrap", "anchor"],
  "wrap <35 calc weight": ["blendWeights", "wrapUnder35", "calculated"],
  "wrap 35 calc weight": ["blendWeights", "wrapUnder35", "calculated"],
  "wrap <35 anchor weight": ["blendWeights", "wrapUnder35", "anchor"],
  "wrap 35 anchor weight": ["blendWeights", "wrapUnder35", "anchor"],
  "wrap <70 calc weight": ["blendWeights", "wrapUnder70", "calculated"],
  "wrap <70 anchor weight": ["blendWeights", "wrapUnder70", "anchor"],
  "wrap 70%+ calc weight": ["blendWeights", "wrapFull", "calculated"],
  "wrap 70+ calc weight": ["blendWeights", "wrapFull", "calculated"],
  "wrap 70%+ anchor weight": ["blendWeights", "wrapFull", "anchor"],
  "wrap 70+ anchor weight": ["blendWeights", "wrapFull", "anchor"],
  "absolute minimum": ["minPrice"],
  "any wrap minimum": ["minAnyWrapPrice"],
  "partial wrap minimum": ["minPartialWrapPrice"],
  "full wrap minimum": ["minFullWrapPrice"]
};

function normalizeSmartPriceLabel(label = "") {
  return String(label)
    .toLowerCase()
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*<\s*/g, " <")
    .replace(/\s*>\s*/g, " >")
    .trim();
}

function setNestedPricingValue(settings, path, value) {
  const next = { ...settings };
  let target = next;
  path.slice(0, -1).forEach((key) => {
    target[key] = { ...target[key] };
    target = target[key];
  });
  target[path[path.length - 1]] = value;
  return next;
}

function getStoredVehiclePriceTrainingBank() {
  if (typeof window === "undefined") return [];
  try {
    const storedBank = window.localStorage.getItem(VEHICLE_PRICE_TRAINING_BANK_STORAGE_KEY);
    const parsedBank = JSON.parse(storedBank || "[]");
    return Array.isArray(parsedBank) ? parsedBank : [];
  } catch (error) {
    return [];
  }
}

function saveVehiclePriceTrainingBank(bank) {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(VEHICLE_PRICE_TRAINING_BANK_STORAGE_KEY, JSON.stringify(bank));
  } catch (error) {
    // Training examples still work for the current session if storage is unavailable.
  }
}

function getSmartPriceImportSection(line = "") {
  const normalizedLine = normalizeSmartPriceLabel(line);
  if (normalizedLine.includes("medium van")) return "medium";
  if (normalizedLine.includes("large van")) return "large";
  return "";
}

function getTemplateSmartPriceSection(template) {
  return normalizeSmartPriceLabel(template?.sizeName).includes("large") ? "large" : "medium";
}

function parseSmartPriceImport(text, template) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const hasSectionHeaders = lines.some((line) => getSmartPriceImportSection(line));
  const targetSection = getTemplateSmartPriceSection(template);
  let activeSection = hasSectionHeaders ? "" : targetSection;
  const updates = [];
  const unknown = [];

  lines.forEach((line) => {
    const nextSection = getSmartPriceImportSection(line);
    if (nextSection) {
      activeSection = nextSection;
      return;
    }
    if (hasSectionHeaders && activeSection !== targetSection) return;

    const match = line.match(/^(.+?)\s+(-?£?[\d,.]+)\s*$/i);
    if (!match) return;

    const label = normalizeSmartPriceLabel(match[1].replace(/:$/, ""));
    const path = SMART_PRICE_IMPORT_FIELDS[label];
    const value = Number(match[2].replace(/[£,]/g, ""));
    if (!path || !Number.isFinite(value)) {
      unknown.push(match[1].trim());
      return;
    }
    updates.push({ label: match[1].trim(), path, value });
  });

  return { updates, unknown, targetSection };
}

function mergeVehiclePricingSettings(settings = {}) {
  return {
    ...VEHICLE_GRAPHICS_PRICING,
    ...settings,
    sectionFactors: { ...VEHICLE_GRAPHICS_PRICING.sectionFactors, ...settings.sectionFactors },
    difficultyFactors: { ...VEHICLE_GRAPHICS_PRICING.difficultyFactors, ...settings.difficultyFactors },
    marketAnchors: { ...VEHICLE_GRAPHICS_PRICING.marketAnchors, ...settings.marketAnchors },
    materialMultipliers: { ...VEHICLE_GRAPHICS_PRICING.materialMultipliers, ...settings.materialMultipliers },
    blendWeights: {
      noWrap: { ...VEHICLE_GRAPHICS_PRICING.blendWeights.noWrap, ...settings.blendWeights?.noWrap },
      wrapUnder35: { ...VEHICLE_GRAPHICS_PRICING.blendWeights.wrapUnder35, ...settings.blendWeights?.wrapUnder35 },
      wrapUnder70: { ...VEHICLE_GRAPHICS_PRICING.blendWeights.wrapUnder70, ...settings.blendWeights?.wrapUnder70 },
      wrapFull: { ...VEHICLE_GRAPHICS_PRICING.blendWeights.wrapFull, ...settings.blendWeights?.wrapFull }
    }
  };
}

function looksLikeVehiclePricingSettings(value) {
  return Boolean(
    value &&
      typeof value === "object" &&
      ("standardVinylRate" in value ||
        "wrapRateStart" in value ||
        "labourSellRate" in value ||
        "marketAnchors" in value ||
        "blendWeights" in value)
  );
}

function getTemplatePricingDefaultsVersion(template) {
  return template.pricingDefaultsVersion || 0;
}

function getDefaultVehiclePricingSettingsByTemplate() {
  return Object.fromEntries(
    VEHICLE_TEMPLATE_OPTIONS.map((template) => [
      template.id,
      {
        ...mergeVehiclePricingSettings(template.pricingSettings || {}),
        __defaultsVersion: getTemplatePricingDefaultsVersion(template)
      }
    ])
  );
}

function normalizeVehiclePricingSettingsByTemplate(settingsByTemplate = {}) {
  return Object.fromEntries(
    VEHICLE_TEMPLATE_OPTIONS.map((template) => {
      const storedTemplateSettings = settingsByTemplate?.[template.id] || {};
      const defaultsVersion = getTemplatePricingDefaultsVersion(template);
      const storedVersion = Number(storedTemplateSettings.__defaultsVersion || 0);
      const shouldRefreshTemplateDefaults = defaultsVersion > storedVersion;
      const mergedSettings = shouldRefreshTemplateDefaults
        ? mergeVehiclePricingSettings(template.pricingSettings || {})
        : mergeVehiclePricingSettings(storedTemplateSettings || template.pricingSettings || {});
      return [
        template.id,
        {
          ...mergedSettings,
          __defaultsVersion: defaultsVersion
        }
      ];
    })
  );
}

function getStoredVehiclePricingSettingsByTemplate() {
  const defaultSettings = getDefaultVehiclePricingSettingsByTemplate();
  if (typeof window === "undefined") return defaultSettings;
  try {
    const storedSettings = window.localStorage.getItem(VEHICLE_PRICING_STORAGE_KEY);
    if (!storedSettings) return defaultSettings;
    const parsedSettings = JSON.parse(storedSettings);

    if (looksLikeVehiclePricingSettings(parsedSettings)) return defaultSettings;

    return normalizeVehiclePricingSettingsByTemplate(parsedSettings);
  } catch (error) {
    return defaultSettings;
  }
}

function saveVehiclePricingSettingsByTemplate(settingsByTemplate) {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(VEHICLE_PRICING_STORAGE_KEY, JSON.stringify(settingsByTemplate));
  } catch (error) {
    // The saved values still apply for this browser session if storage is unavailable.
  }
}

function decodeSvgLayerId(value = "") {
  return String(value || "").replace(/_x([0-9a-fA-F]{4})_/g, (_, hex) =>
    String.fromCharCode(parseInt(hex, 16))
  );
}

function getScaleReferenceMm(value = "") {
  const match = decodeSvgLayerId(value).match(/(\d+(?:\.\d+)?)\s*mm/i);
  return match ? Number(match[1]) : 0;
}

function normalizeVehicleScaleFactor(computedScale, fallbackScale, artworkScale = 1) {
  const computed = Number(computedScale);
  const fallback = Number(fallbackScale);
  if (!Number.isFinite(computed) || computed <= 0) return fallback;
  if (!Number.isFinite(fallback) || fallback <= 0) return computed;

  const scaleAdjustment = Number(artworkScale) > 0 && Number(artworkScale) < 1 ? 1 / Number(artworkScale) : 1;
  const adjustedForArtworkScale = computed * scaleAdjustment;

  if (
    computed < fallback * 0.5 &&
    adjustedForArtworkScale >= fallback * 0.5 &&
    adjustedForArtworkScale <= fallback * 1.5
  ) {
    return adjustedForArtworkScale;
  }

  if (computed > fallback * 2 && computed / scaleAdjustment >= fallback * 0.5 && computed / scaleAdjustment <= fallback * 1.5) {
    return computed / scaleAdjustment;
  }

  return computed;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeSvgText(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function VinylEstimatorPage({ currentUser, onLogout, notifications }) {
  const pricingEditable = canEditVanEstimator(currentUser);
  const [svgMarkup, setSvgMarkup] = useState("");
  const [artBoardMarkup, setArtBoardMarkup] = useState("");
  const [svgError, setSvgError] = useState("");
  const [shapes, setShapes] = useState([]);
  const [pdfExportOpen, setPdfExportOpen] = useState(false);
  const [pdfExporting, setPdfExporting] = useState(false);
  const [pdfExportError, setPdfExportError] = useState("");
  const [pdfExportForm, setPdfExportForm] = useState({ customerName: "", notes: "" });
  const [selectedTemplateId, setSelectedTemplateId] = useState(VAN_ESTIMATOR_TEMPLATE.id);
  const selectedTemplate =
    VEHICLE_TEMPLATE_OPTIONS.find((template) => template.id === selectedTemplateId) || VAN_ESTIMATOR_TEMPLATE;
  const [pricingSettingsByTemplate, setPricingSettingsByTemplate] = useState(getStoredVehiclePricingSettingsByTemplate);
  const pricingSettings = useMemo(
    () => mergeVehiclePricingSettings(pricingSettingsByTemplate[selectedTemplate.id] || selectedTemplate.pricingSettings || {}),
    [pricingSettingsByTemplate, selectedTemplate]
  );
  const [pricingDraftSettings, setPricingDraftSettings] = useState(() => pricingSettings);
  const [pricingSettingsOpen, setPricingSettingsOpen] = useState(false);
  const [smartPriceImportOpen, setSmartPriceImportOpen] = useState(false);
  const [smartPriceImportText, setSmartPriceImportText] = useState("");
  const [smartPriceImportStatus, setSmartPriceImportStatus] = useState("");
  const [pricingSyncStatus, setPricingSyncStatus] = useState("");
  const [priceTrainingBank, setPriceTrainingBank] = useState(getStoredVehiclePriceTrainingBank);
  const [trainingTargetPrice, setTrainingTargetPrice] = useState("");
  const [trainingStatus, setTrainingStatus] = useState("");
  const [trainingSuggestion, setTrainingSuggestion] = useState(null);
  const [activeScaleFactor, setActiveScaleFactor] = useState(selectedTemplate.scaleFactor);
  const [drawMode, setDrawMode] = useState("rectangle");
  const [materialMode, setMaterialMode] = useState("standard");
  const [drawingRect, setDrawingRect] = useState(null);
  const [drawStart, setDrawStart] = useState(null);
  const [polygonPoints, setPolygonPoints] = useState([]);
  const [polygonPreviewPoint, setPolygonPreviewPoint] = useState(null);
  const [lassoPoints, setLassoPoints] = useState([]);
  const [editDrag, setEditDrag] = useState(null);
  const [vehicleClipPathsD, setVehicleClipPathsD] = useState([]);
  const [vehicleEdgePathsD, setVehicleEdgePathsD] = useState([]);
  const inlineSvgRef = useRef(null);
  const overlaySvgRef = useRef(null);
  const vehicleBodyPathsRef = useRef([]);
  const wrapLinesRef = useRef([]);
  const vehiclePricingSeededRef = useRef(false);

  const selectedTemplateTrainingBank = useMemo(
    () => priceTrainingBank.filter((sample) => sample.templateId === selectedTemplate.id),
    [priceTrainingBank, selectedTemplate.id]
  );

  useEffect(() => {
    let active = true;
    fetch("/vans/art-board.svg")
      .then((response) => {
        if (!response.ok) throw new Error("Could not load the art board SVG.");
        return response.text();
      })
      .then((text) => {
        if (active) setArtBoardMarkup(text);
      })
      .catch((error) => {
        console.error(error);
      });

    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    let active = true;
    setSvgMarkup("");
    setSvgError("");
    setShapes([]);
    setDrawingRect(null);
    setDrawStart(null);
    setPolygonPoints([]);
    setPolygonPreviewPoint(null);
    setLassoPoints([]);
    setEditDrag(null);
    setVehicleClipPathsD([]);
    setVehicleEdgePathsD([]);
    setActiveScaleFactor(selectedTemplate.scaleFactor);
    vehicleBodyPathsRef.current = [];
    wrapLinesRef.current = [];

    fetch(selectedTemplate.src)
      .then((response) => {
        if (!response.ok) throw new Error("Could not load the vehicle SVG.");
        return response.text();
      })
      .then((text) => {
        if (!active) return;
        setSvgMarkup(text);
        setSvgError("");
      })
      .catch((error) => {
        console.error(error);
        if (active) setSvgError(error.message || "Could not load the van SVG.");
      });
    return () => {
      active = false;
    };
  }, [selectedTemplate]);

  useEffect(() => {
    setPricingDraftSettings(pricingSettings);
    setTrainingSuggestion(null);
    setTrainingStatus("");
  }, [selectedTemplate.id]);

  useEffect(() => {
    let active = true;

    async function loadSharedPricingSettings() {
      try {
        const response = await fetch("/api/vehicle-pricing/settings");
        const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload.error || "Could not load shared vehicle pricing.");
        if (!active) return;

        if (payload.settingsByTemplate && typeof payload.settingsByTemplate === "object") {
          const nextSettingsByTemplate = normalizeVehiclePricingSettingsByTemplate(payload.settingsByTemplate);
          setPricingSettingsByTemplate(nextSettingsByTemplate);
          saveVehiclePricingSettingsByTemplate(nextSettingsByTemplate);
          setPricingSyncStatus("Shared pricing loaded.");
          return;
        }

        if (pricingEditable && !vehiclePricingSeededRef.current) {
          vehiclePricingSeededRef.current = true;
          const seedSettings = normalizeVehiclePricingSettingsByTemplate(pricingSettingsByTemplate);
          const saveResponse = await fetch("/api/vehicle-pricing/settings", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ settingsByTemplate: seedSettings })
          });
          const savePayload = await saveResponse.json().catch(() => ({}));
          if (!saveResponse.ok) throw new Error(savePayload.error || "Could not publish vehicle pricing.");
          if (!active) return;
          const nextSettingsByTemplate = normalizeVehiclePricingSettingsByTemplate(savePayload.settingsByTemplate || seedSettings);
          setPricingSettingsByTemplate(nextSettingsByTemplate);
          saveVehiclePricingSettingsByTemplate(nextSettingsByTemplate);
          setPricingSyncStatus("Shared pricing published.");
        }
      } catch (error) {
        console.error(error);
        if (active) setPricingSyncStatus("Using this browser's saved pricing.");
      }
    }

    loadSharedPricingSettings();
    return () => {
      active = false;
    };
  }, [pricingEditable]);

  useEffect(() => {
    saveVehiclePriceTrainingBank(priceTrainingBank);
  }, [priceTrainingBank]);

  useEffect(() => {
    if (!svgMarkup || !inlineSvgRef.current) return;
    const svg = inlineSvgRef.current.querySelector("svg");
    if (svg) {
      svg.setAttribute("width", "100%");
      svg.removeAttribute("height");
      svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
      svg.classList.add("van-template-svg");
    }

    const tyreReference = selectedTemplate.tyreReferenceSelector
      ? Array.from(inlineSvgRef.current.querySelectorAll(selectedTemplate.tyreReferenceSelector))
          .map((element) => {
            try {
              const box = element.getBBox();
              const diameterUnits = Math.max(box.width || 0, box.height || 0);
              const aspectRatio = box.width && box.height ? box.width / box.height : 0;
              return diameterUnits > 50 && aspectRatio > 0.75 && aspectRatio < 1.33 ? { element, diameterUnits } : null;
            } catch (error) {
              return null;
            }
          })
          .filter(Boolean)
          .sort((left, right) => right.diameterUnits - left.diameterUnits)[0]
      : null;

    if (tyreReference && selectedTemplate.tyreReferenceDiameterMm) {
      setActiveScaleFactor(selectedTemplate.tyreReferenceDiameterMm / tyreReference.diameterUnits);
    } else {
      const scaleReferenceLayer = Array.from(inlineSvgRef.current.querySelectorAll("g[id]"))
        .map((element) => {
          const referenceMm = getScaleReferenceMm(element.id);
          return referenceMm ? { element, referenceMm } : null;
        })
        .filter(Boolean)
        .sort((left, right) => right.referenceMm - left.referenceMm)[0];
      if (scaleReferenceLayer) {
        try {
          const box = scaleReferenceLayer.element.getBBox();
          const referenceUnits = Math.max(box.width || 0, box.height || 0);
          if (referenceUnits > 0) {
            setActiveScaleFactor(
              normalizeVehicleScaleFactor(
                scaleReferenceLayer.referenceMm / referenceUnits,
                selectedTemplate.scaleFactor,
                selectedTemplate.artworkScale
              )
            );
          }
        } catch (error) {
          setActiveScaleFactor(selectedTemplate.scaleFactor);
        }
      } else {
        setActiveScaleFactor(selectedTemplate.scaleFactor);
      }
    }

    const artworkLayer = inlineSvgRef.current.querySelector("#Artwork");
    const edgeLayer = inlineSvgRef.current.querySelector("#Van_Edges");
    const vehicleBodyLayer = artworkLayer || edgeLayer;
    const vehicleBodyPaths = Array.from(vehicleBodyLayer?.querySelectorAll("path") || [])
      .map((element) => {
        try {
          const box = element.getBBox();
          return { element, area: box.width * box.height };
        } catch (error) {
          return null;
        }
      })
      .filter(Boolean)
      .filter((entry) => entry.area > 100000)
      .map((entry) => entry.element);
    vehicleBodyPathsRef.current = vehicleBodyPaths;
    setVehicleClipPathsD(vehicleBodyPaths.map((element) => element.getAttribute("d")).filter(Boolean));
    setVehicleEdgePathsD(
      Array.from((edgeLayer || vehicleBodyLayer)?.querySelectorAll("path") || [])
        .map((element) => element.getAttribute("d"))
        .filter(Boolean)
    );

    const wrapLayer = inlineSvgRef.current.querySelector("#Wrap_Film_Lines");
    if (!wrapLayer) {
      wrapLinesRef.current = [];
      return;
    }

    wrapLinesRef.current = Array.from(wrapLayer.querySelectorAll("path,line,polyline,polygon"))
      .filter((element) => {
        const styles = window.getComputedStyle(element);
        const fill = styles.fill || element.getAttribute("fill") || "";
        const stroke = styles.stroke || element.getAttribute("stroke") || "";
        return fill === "none" && stroke !== "none";
      })
      .map((element) => {
        try {
          const box = element.getBBox();
          const length = typeof element.getTotalLength === "function" ? element.getTotalLength() : 0;
          const sampleCount = Math.max(2, Math.ceil(length / 8));
          const points =
            length > 0
              ? Array.from({ length: sampleCount + 1 }, (_, index) => {
                  const point = element.getPointAtLength((length * index) / sampleCount);
                  return { x: point.x, y: point.y };
                })
              : [];
          return {
            box: {
              x: box.x,
              y: box.y,
              width: box.width,
              height: box.height
            },
            points
          };
        } catch (error) {
          return null;
        }
      })
      .filter((line) => line?.points?.length);
  }, [svgMarkup, selectedTemplate]);

  function getPointerPoint(event) {
    const svg = overlaySvgRef.current;
    if (!svg) return null;
    const point = svg.createSVGPoint();
    point.x = event.clientX;
    point.y = event.clientY;
    const matrix = svg.getScreenCTM();
    if (!matrix) return null;
    return point.matrixTransform(matrix.inverse());
  }

  function normalizeRect(start, end) {
    const x = Math.min(start.x, end.x);
    const y = Math.min(start.y, end.y);
    const width = Math.abs(end.x - start.x);
    const height = Math.abs(end.y - start.y);
    return { x, y, width, height };
  }

  function getPolygonBounds(points) {
    const xValues = points.map((point) => point.x);
    const yValues = points.map((point) => point.y);
    const x = Math.min(...xValues);
    const y = Math.min(...yValues);
    return {
      x,
      y,
      width: Math.max(...xValues) - x,
      height: Math.max(...yValues) - y
    };
  }

  function pointsToSvg(points) {
    return points.map((point) => `${point.x},${point.y}`).join(" ");
  }

  function getShapeClipId(shapeId) {
    return `vinyl-shape-clip-${String(shapeId).replace(/[^a-zA-Z0-9_-]/g, "-")}`;
  }

  function getRectanglePoints(rect) {
    return [
      { corner: "top-left", x: rect.x, y: rect.y },
      { corner: "top-right", x: rect.x + rect.width, y: rect.y },
      { corner: "bottom-right", x: rect.x + rect.width, y: rect.y + rect.height },
      { corner: "bottom-left", x: rect.x, y: rect.y + rect.height }
    ];
  }

  function getVehicleBodyBounds() {
    const boxes = [];
    vehicleBodyPathsRef.current.forEach((path) => {
      try {
        const box = path.getBBox();
        if (box?.width && box?.height) boxes.push(box);
      } catch (error) {
        // Ignore paths that are not ready for geometry calculations.
      }
    });
    if (boxes.length) {
      const minX = Math.min(...boxes.map((box) => box.x));
      const minY = Math.min(...boxes.map((box) => box.y));
      const maxX = Math.max(...boxes.map((box) => box.x + box.width));
      const maxY = Math.max(...boxes.map((box) => box.y + box.height));
      return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
    }

    try {
      const svg = inlineSvgRef.current?.querySelector("svg");
      const box = svg?.getBBox();
      if (box?.width && box?.height) return { x: box.x, y: box.y, width: box.width, height: box.height };
    } catch (error) {
      // Fall back to the full drawing viewBox if the SVG is not ready yet.
    }
    return selectedTemplate.viewBox;
  }

  function rectsIntersect(left, right) {
    return (
      left.x < right.x + right.width &&
      left.x + left.width > right.x &&
      left.y < right.y + right.height &&
      left.y + left.height > right.y
    );
  }

  function isPointInPolygon(point, polygon) {
    let inside = false;
    for (let index = 0, previousIndex = polygon.length - 1; index < polygon.length; previousIndex = index++) {
      const current = polygon[index];
      const previous = polygon[previousIndex];
      const crossesY = current.y > point.y !== previous.y > point.y;
      if (!crossesY) continue;
      const crossingX = ((previous.x - current.x) * (point.y - current.y)) / (previous.y - current.y) + current.x;
      if (point.x < crossingX) inside = !inside;
    }
    return inside;
  }

  function isPointInVehicleBody(point) {
    const svg = inlineSvgRef.current?.querySelector("svg");
    if (!vehicleBodyPathsRef.current.length || !svg) return true;

    try {
      const svgPoint = svg.createSVGPoint();
      svgPoint.x = point.x;
      svgPoint.y = point.y;
      if (vehicleBodyPathsRef.current.some((path) => typeof path.isPointInFill === "function" && path.isPointInFill(svgPoint))) {
        return true;
      }
    } catch (error) {
      // Older engines can miss SVG geometry helpers; the bbox fallback still prevents wild off-vehicle areas.
    }

    return vehicleBodyPathsRef.current.some((path) => {
      try {
        const bounds = path.getBBox();
        return (
          point.x >= bounds.x &&
          point.x <= bounds.x + bounds.width &&
          point.y >= bounds.y &&
          point.y <= bounds.y + bounds.height
        );
      } catch (error) {
        return false;
      }
    });
  }

  function getVehicleClipRatio(area) {
    if (!vehicleBodyPathsRef.current.length) return 1;
    const bounds = area.bounds || area;
    if (!bounds?.width || !bounds?.height) return 1;

    const columns = Math.max(4, Math.min(32, Math.ceil(bounds.width / 24)));
    const rows = Math.max(4, Math.min(32, Math.ceil(bounds.height / 24)));
    let insideShape = 0;
    let insideVehicle = 0;

    for (let row = 0; row < rows; row += 1) {
      for (let column = 0; column < columns; column += 1) {
        const point = {
          x: bounds.x + ((column + 0.5) * bounds.width) / columns,
          y: bounds.y + ((row + 0.5) * bounds.height) / rows
        };
        if (area.points?.length && !isPointInPolygon(point, area.points)) continue;
        insideShape += 1;
        if (isPointInVehicleBody(point)) insideVehicle += 1;
      }
    }

    return insideShape > 0 ? insideVehicle / insideShape : 0;
  }

  function isWrapFilmArea(area) {
    const detectionPadding = 2;
    const rect = area.bounds || area;
    const expandedRect = {
      x: rect.x - detectionPadding,
      y: rect.y - detectionPadding,
      width: rect.width + detectionPadding * 2,
      height: rect.height + detectionPadding * 2
    };

    return wrapLinesRef.current.some((line) => {
      if (!rectsIntersect(expandedRect, line.box)) return false;
      return line.points.some((point) => {
        if (
          point.x < expandedRect.x ||
          point.x > expandedRect.x + expandedRect.width ||
          point.y < expandedRect.y ||
          point.y > expandedRect.y + expandedRect.height
        ) {
          return false;
        }
        return area.points?.length ? isPointInPolygon(point, area.points) : true;
      });
    });
  }

  function getVehicleZoneMetadata(area) {
    const wrapRequired = isWrapFilmArea(area);
    const bounds = area.bounds || area;
    const bodyBounds = getVehicleBodyBounds();
    const centerX = bounds.x + bounds.width / 2;
    const frontThreshold = bodyBounds.x + bodyBounds.width * 0.62;
    const rearThreshold = bodyBounds.x + bodyBounds.width * 0.38;

    if (wrapRequired) {
      const installGroup = centerX >= frontThreshold ? "front_end" : centerX <= rearThreshold ? "rear_half" : "main_face";
      return {
        material_type: "wrap_film",
        surface_type: "normal_wrap_curve",
        complexity_factor: 1.15,
        install_group: installGroup
      };
    }

    return {
      material_type: "standard_vinyl",
      surface_type: "flat",
      complexity_factor: 1,
      install_group: "standard_panel"
    };
  }

  function getForcedStandardZoneMetadata(materialVariant = "standard") {
    return {
      material_type: "standard_vinyl",
      surface_type: "flat",
      complexity_factor: 1,
      install_group: materialVariant === "standard" ? "standard_panel" : materialVariant
    };
  }

  function getForcedWrapZoneMetadata(area) {
    const bounds = area.bounds || area;
    const bodyBounds = getVehicleBodyBounds();
    const centerX = bounds.x + bounds.width / 2;
    const frontThreshold = bodyBounds.x + bodyBounds.width * 0.62;
    const rearThreshold = bodyBounds.x + bodyBounds.width * 0.38;
    const installGroup = centerX >= frontThreshold ? "front_end" : centerX <= rearThreshold ? "rear_half" : "main_face";

    return {
      material_type: "wrap_film",
      surface_type: "normal_wrap_curve",
      complexity_factor: 1.15,
      install_group: installGroup
    };
  }

  function getShapeMaterialVariant(shape) {
    return ["contra", "reflective"].includes(shape.materialVariant) ? shape.materialVariant : "standard";
  }

  function interpolatePricingPoint(points, value) {
    if (value <= points[0].coverage) return points[0].value;
    for (let index = 1; index < points.length; index += 1) {
      const previous = points[index - 1];
      const next = points[index];
      if (value > next.coverage) continue;
      const progress = (value - previous.coverage) / (next.coverage - previous.coverage || 1);
      return previous.value + (next.value - previous.value) * progress;
    }
    return points[points.length - 1].value;
  }

  function getFormulaWrapRate(wrapCoverage, settings) {
    return Math.max(settings.wrapRateFloor, settings.wrapRateStart - settings.wrapRateTaper * Math.sqrt(Math.max(0, wrapCoverage)));
  }

  function getFormulaWrapLabourHours(wrapCoverage, settings) {
    return Math.max(
      settings.wrapLabourFloorHoursPerM2,
      settings.wrapLabourStartHoursPerM2 - settings.wrapLabourTaper * Math.sqrt(Math.max(0, wrapCoverage))
    );
  }

  function getMarketAnchor(coverage, settings) {
    return interpolatePricingPoint(
      [
        { coverage: 0, value: settings.marketAnchors.c0 },
        { coverage: 0.05, value: settings.marketAnchors.c05 },
        { coverage: 0.1, value: settings.marketAnchors.c10 },
        { coverage: 0.15, value: settings.marketAnchors.c15 },
        { coverage: 0.22, value: settings.marketAnchors.c22 },
        { coverage: 0.35, value: settings.marketAnchors.c35 },
        { coverage: 0.55, value: settings.marketAnchors.c55 },
        { coverage: 0.85, value: settings.marketAnchors.c85 },
        { coverage: 1, value: settings.marketAnchors.c100 }
      ],
      coverage
    );
  }

  function getBlendWeights(coverage, wrapArea, settings) {
    if (wrapArea <= 0) return settings.blendWeights.noWrap;
    if (coverage < 0.35) return settings.blendWeights.wrapUnder35;
    if (coverage < 0.7) return settings.blendWeights.wrapUnder70;
    return settings.blendWeights.wrapFull;
  }

  function clampNumber(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function getVehicleDrawableAreaM2() {
    if (!vehicleBodyPathsRef.current.length) {
      return (
        (selectedTemplate.viewBox.width *
          selectedTemplate.viewBox.height *
          activeScaleFactor *
          activeScaleFactor) /
        1000000
      );
    }

    const bounds = getVehicleBodyBounds();
    const columns = 90;
    const rows = 52;
    let insideVehicle = 0;

    for (let row = 0; row < rows; row += 1) {
      for (let column = 0; column < columns; column += 1) {
        const point = {
          x: bounds.x + ((column + 0.5) * bounds.width) / columns,
          y: bounds.y + ((row + 0.5) * bounds.height) / rows
        };
        if (isPointInVehicleBody(point)) insideVehicle += 1;
      }
    }

    const unitsArea = bounds.width * bounds.height * (insideVehicle / (columns * rows));
    const scaleFactor = activeScaleFactor;
    return (unitsArea * scaleFactor * scaleFactor) / 1000000;
  }

  function getShapeCenter(shape) {
    const bounds = shape.bounds || shape;
    return {
      x: bounds.x + bounds.width / 2,
      y: bounds.y + bounds.height / 2
    };
  }

  function countConnectedShapeSections(sectionShapes) {
    if (!sectionShapes.length) return 0;
    const visited = new Set();
    const expansion = 12;

    function expandedBounds(shape) {
      const bounds = shape.bounds || shape;
      return {
        x: bounds.x - expansion,
        y: bounds.y - expansion,
        width: bounds.width + expansion * 2,
        height: bounds.height + expansion * 2
      };
    }

    let sections = 0;
    sectionShapes.forEach((shape, index) => {
      if (visited.has(index)) return;
      sections += 1;
      const queue = [index];
      visited.add(index);

      while (queue.length) {
        const currentIndex = queue.shift();
        const currentBounds = expandedBounds(sectionShapes[currentIndex]);
        sectionShapes.forEach((candidate, candidateIndex) => {
          if (visited.has(candidateIndex)) return;
          if (!rectsIntersect(currentBounds, expandedBounds(candidate))) return;
          visited.add(candidateIndex);
          queue.push(candidateIndex);
        });
      }
    });

    return sections;
  }

  function getWrapSectionFactor(sectionCount, settings) {
    if (sectionCount <= 1) return settings.sectionFactors.one;
    if (sectionCount <= 3) return settings.sectionFactors.twoToThree;
    if (sectionCount <= 5) return settings.sectionFactors.fourToFive;
    return settings.sectionFactors.moreThanFive;
  }

  function getWeightedDifficultyFactor(wrapShapes, wrapArea, settings) {
    if (!wrapShapes.length || wrapArea <= 0) return 1;
    const weightedDifficulty = wrapShapes.reduce((sum, shape) => {
      const surfaceType = shape.zoneMetadata?.surface_type || "normal_wrap_curve";
      const factor = settings.difficultyFactors[surfaceType] || shape.zoneMetadata?.complexity_factor || 1;
      return sum + shape.areaM2 * factor;
    }, 0);
    return weightedDifficulty / wrapArea;
  }

  function clearTextSelection() {
    if (typeof window !== "undefined") window.getSelection?.()?.removeAllRanges();
  }

  function getRectAreaM2(rect) {
    const widthMm = rect.width * activeScaleFactor;
    const heightMm = rect.height * activeScaleFactor;
    return (widthMm / 1000) * (heightMm / 1000);
  }

  function getPolygonAreaM2(points) {
    const areaUnits =
      Math.abs(
        points.reduce((sum, point, index) => {
          const nextPoint = points[(index + 1) % points.length];
          return sum + point.x * nextPoint.y - nextPoint.x * point.y;
        }, 0)
      ) / 2;
    const scaleFactor = activeScaleFactor;
    return (areaUnits * scaleFactor * scaleFactor) / 1000000;
  }

  function refreshShapeMetrics(shape) {
    const materialVariant = getShapeMaterialVariant(shape);
    if (shape.type === "polygon") {
      const bounds = getPolygonBounds(shape.points);
      const zoneMetadata =
        shape.materialOverride === "wrap"
          ? getForcedWrapZoneMetadata({ points: shape.points, bounds })
          : shape.materialOverride === "standard"
          ? getForcedStandardZoneMetadata("standard")
          : materialVariant === "standard"
          ? getVehicleZoneMetadata({ points: shape.points, bounds })
          : getForcedStandardZoneMetadata(materialVariant);
      const rawAreaM2 = getPolygonAreaM2(shape.points);
      const clipRatio = getVehicleClipRatio({ points: shape.points, bounds });
      return {
        ...shape,
        materialVariant,
        bounds,
        width: bounds.width,
        height: bounds.height,
        materialOverride: shape.materialOverride || "",
        rawAreaM2,
        clipRatio,
        areaM2: rawAreaM2 * clipRatio,
        zoneMetadata,
        isWrapFilm: zoneMetadata.material_type === "wrap_film"
      };
    }

    const rect = {
      x: shape.x,
      y: shape.y,
      width: shape.width,
      height: shape.height
    };
    const zoneMetadata =
      shape.materialOverride === "wrap"
        ? getForcedWrapZoneMetadata(rect)
        : shape.materialOverride === "standard"
        ? getForcedStandardZoneMetadata("standard")
        : materialVariant === "standard"
        ? getVehicleZoneMetadata(rect)
        : getForcedStandardZoneMetadata(materialVariant);
    const rawAreaM2 = getRectAreaM2(rect);
    const clipRatio = getVehicleClipRatio(rect);

    return {
      ...shape,
      materialVariant,
      materialOverride: shape.materialOverride || "",
      bounds: rect,
      rawAreaM2,
      clipRatio,
      areaM2: rawAreaM2 * clipRatio,
      zoneMetadata,
      isWrapFilm: zoneMetadata.material_type === "wrap_film"
    };
  }

  useEffect(() => {
    setShapes((current) => current.map((shape) => refreshShapeMetrics(shape)));
  }, [activeScaleFactor]);

  function updateShapeCorner(shape, dragState, point) {
    if (shape.type === "polygon") {
      const points = shape.points.map((entry, index) => (index === dragState.pointIndex ? point : entry));
      return refreshShapeMetrics({ ...shape, points });
    }

    const anchor = dragState.anchor;
    const rect = normalizeRect(anchor, point);
    return refreshShapeMetrics({
      ...shape,
      ...rect
    });
  }

  function finishPolygon() {
    if (polygonPoints.length < 3) return;
    const shape = createPolygonShape(polygonPoints);
    if (shape) setShapes((current) => [...current, shape]);
    setPolygonPoints([]);
    setPolygonPreviewPoint(null);
  }

  function getDistanceBetweenPoints(first, second) {
    return Math.hypot(first.x - second.x, first.y - second.y);
  }

  function shouldClosePolygon(point, points) {
    if (points.length < 3) return false;
    const firstPoint = points[0];
    return getDistanceBetweenPoints(point, firstPoint) <= 28;
  }

  function simplifyFreehandPoints(points) {
    if (points.length <= 2) return points;
    return points.reduce((result, point, index) => {
      if (index === 0 || index === points.length - 1) return [...result, point];
      const previousPoint = result[result.length - 1];
      return getDistanceBetweenPoints(point, previousPoint) >= 10 ? [...result, point] : result;
    }, []);
  }

  function createPolygonShape(points, idPrefix = "vinyl-poly") {
    if (points.length < 3) return null;
    const bounds = getPolygonBounds(points);
    const shape = refreshShapeMetrics({
      id: `${idPrefix}-${Date.now()}-${Math.round(bounds.x)}-${Math.round(bounds.y)}`,
      type: "polygon",
      materialVariant: materialMode,
      points,
      bounds,
      width: bounds.width,
      height: bounds.height
    });
    return shape.areaM2 > 0.001 ? shape : null;
  }

  function startDrawing(event) {
    if (event.button !== 0) return;
    if (editDrag) return;
    if (event.target.closest?.(".vinyl-canvas-toolbar, .vinyl-material-toolbar")) return;
    clearTextSelection();
    const point = getPointerPoint(event);
    if (!point) return;
    if (drawMode === "polygon") {
      if (shouldClosePolygon(point, polygonPoints)) {
        finishPolygon();
        return;
      }
      setPolygonPoints((current) => [...current, point]);
      setPolygonPreviewPoint(null);
      return;
    }
    if (drawMode === "lasso") {
      event.currentTarget.setPointerCapture(event.pointerId);
      setLassoPoints([point]);
      setDrawStart(null);
      setDrawingRect(null);
      return;
    }
    event.currentTarget.setPointerCapture(event.pointerId);
    setDrawStart(point);
    setDrawingRect({ x: point.x, y: point.y, width: 0, height: 0 });
  }

  function updateDrawing(event) {
    const point = getPointerPoint(event);
    if (!point) return;
    if (editDrag) {
      setShapes((current) =>
        current.map((shape) => (shape.id === editDrag.shapeId ? updateShapeCorner(shape, editDrag, point) : shape))
      );
      return;
    }
    if (drawMode === "polygon") {
      if (polygonPoints.length) setPolygonPreviewPoint(point);
      return;
    }
    if (drawMode === "lasso") {
      if (!lassoPoints.length) return;
      setLassoPoints((current) => {
        const previousPoint = current[current.length - 1];
        if (previousPoint && getDistanceBetweenPoints(point, previousPoint) < 6) return current;
        return [...current, point];
      });
      return;
    }
    if (!drawStart) return;
    setDrawingRect(normalizeRect(drawStart, point));
  }

  function finishDrawing(event) {
    if (editDrag) {
      try {
        event.currentTarget.releasePointerCapture(event.pointerId);
      } catch (error) {
        // Pointer capture may already be released if the pointer leaves the browser chrome.
      }
      setEditDrag(null);
      return;
    }
    if (drawMode === "polygon") return;
    if (drawMode === "lasso") {
      try {
        event.currentTarget.releasePointerCapture(event.pointerId);
      } catch (error) {
        // Pointer capture may already be released if the pointer leaves the browser chrome.
      }
      const releasePoint = getPointerPoint(event);
      const points = simplifyFreehandPoints(releasePoint ? [...lassoPoints, releasePoint] : lassoPoints);
      const bounds = points.length >= 2 ? getPolygonBounds(points) : { width: 0, height: 0 };
      if (points.length < 4 || Math.max(bounds.width, bounds.height) < 8) {
        setLassoPoints([]);
        return;
      }
      const shape = createPolygonShape(points, "vinyl-lasso");
      if (shape) setShapes((current) => [...current, shape]);
      setLassoPoints([]);
      return;
    }
    if (!drawStart || !drawingRect) return;
    try {
      event.currentTarget.releasePointerCapture(event.pointerId);
    } catch (error) {
      // Pointer capture may already be released if the pointer leaves the browser chrome.
    }

    const rect = {
      ...drawingRect,
      id: `vinyl-${Date.now()}-${Math.round(drawingRect.x)}-${Math.round(drawingRect.y)}`
    };
    setDrawStart(null);
    setDrawingRect(null);
    if (rect.width < 4 || rect.height < 4) return;

    const shape = refreshShapeMetrics({
      ...rect,
      type: "rectangle",
      materialVariant: materialMode,
      bounds: rect
    });
    if (shape.areaM2 <= 0.001) return;
    setShapes((current) => [...current, shape]);
  }

  function startShapeCornerDrag(event, shape, cornerOrPointIndex) {
    event.stopPropagation();
    event.preventDefault();
    if (event.button !== 0) return;
    clearTextSelection();
    try {
      event.currentTarget.setPointerCapture(event.pointerId);
    } catch (error) {
      // Some browsers may not allow capture on SVG children; the overlay still receives moves.
    }

    if (shape.type === "polygon") {
      setEditDrag({ shapeId: shape.id, pointIndex: cornerOrPointIndex });
      return;
    }

    const rect = shape.bounds || shape;
    const anchors = {
      "top-left": { x: rect.x + rect.width, y: rect.y + rect.height },
      "top-right": { x: rect.x, y: rect.y + rect.height },
      "bottom-right": { x: rect.x, y: rect.y },
      "bottom-left": { x: rect.x + rect.width, y: rect.y }
    };
    setEditDrag({
      shapeId: shape.id,
      corner: cornerOrPointIndex,
      anchor: anchors[cornerOrPointIndex]
    });
  }

  function deleteShape(event, shapeId) {
    event.stopPropagation();
    event.preventDefault();
    clearTextSelection();
    setShapes((current) => current.filter((shape) => shape.id !== shapeId));
  }

  function toggleShapeMaterial(event, shapeId) {
    event.stopPropagation();
    event.preventDefault();
    clearTextSelection();
    setShapes((current) =>
      current.map((shape) => {
        if (shape.id !== shapeId) return shape;
        const nextOverride = shape.isWrapFilm ? "standard" : "wrap";
        return refreshShapeMetrics({
          ...shape,
          materialVariant: "standard",
          materialOverride: nextOverride
        });
      })
    );
  }

  function selectDrawMode(nextMode) {
    setDrawMode(nextMode);
    setDrawStart(null);
    setDrawingRect(null);
    setPolygonPoints([]);
    setPolygonPreviewPoint(null);
    setLassoPoints([]);
  }

  function undoLastDrawing() {
    if (lassoPoints.length) {
      setLassoPoints([]);
      return;
    }
    if (polygonPoints.length) {
      setPolygonPoints((current) => current.slice(0, -1));
      setPolygonPreviewPoint(null);
      return;
    }
    setShapes((current) => current.slice(0, -1));
  }

  function clearDrawing() {
    setDrawStart(null);
    setDrawingRect(null);
    setPolygonPoints([]);
    setPolygonPreviewPoint(null);
    setLassoPoints([]);
    setShapes([]);
  }

  function calculateEstimateFromMetrics(metrics, settings) {
    const standardPrintArea = Number(metrics.standardPrintArea) || 0;
    const contraArea = Number(metrics.contraArea) || 0;
    const reflectiveArea = Number(metrics.reflectiveArea) || 0;
    const wrapArea = Number(metrics.wrapArea) || 0;
    const standardArea = standardPrintArea + contraArea + reflectiveArea;
    const totalArea = standardArea + wrapArea;
    const vehicleArea = Number(metrics.vehicleArea) || 0;
    const wrapSectionCount = Number(metrics.wrapSectionCount) || 0;
    const wrapSectionFactor = wrapArea > 0 ? getWrapSectionFactor(wrapSectionCount, settings) : 1;
    const weightedDifficulty = Number(metrics.weightedDifficulty) || 1;

    if (totalArea <= 0) return 0;

    function calculateScaledEstimate(scale = 1) {
      const scaledStandardArea = standardArea * scale;
      const scaledStandardMaterialMultiplierArea =
        (standardPrintArea * (settings.materialMultipliers.standard || 1) +
          contraArea * (settings.materialMultipliers.contra || 1) +
          reflectiveArea * (settings.materialMultipliers.reflective || 1)) *
        scale;
      const scaledWrapArea = wrapArea * scale;
      const scaledTotalArea = totalArea * scale;
      const scaledCoverage = vehicleArea > 0 ? scaledTotalArea / vehicleArea : 0;
      const scaledWrapCoverage = vehicleArea > 0 ? scaledWrapArea / vehicleArea : 0;
      const standardSell = scaledStandardMaterialMultiplierArea * settings.standardVinylRate;
      const wrapSell = scaledWrapArea * getFormulaWrapRate(scaledWrapCoverage, settings);
      const standardLabourHours =
        scaledStandardArea <= 0
          ? 0
          : scaledCoverage < 0.15
            ? clampNumber(
                scaledStandardArea * settings.standardSmallHoursPerM2,
                settings.standardSmallMinHours,
                settings.standardSmallMaxHours
              )
            : scaledStandardArea * settings.standardLargeHoursPerM2;
      const wrapLabourHours =
        scaledWrapArea * getFormulaWrapLabourHours(scaledWrapCoverage, settings) * wrapSectionFactor * weightedDifficulty;
      const basePrice = standardSell + wrapSell + (standardLabourHours + wrapLabourHours) * settings.labourSellRate;
      const blend = getBlendWeights(scaledCoverage, scaledWrapArea, settings);
      let estimate = blend.calculated * basePrice + blend.anchor * getMarketAnchor(scaledCoverage, settings);

      estimate = Math.max(estimate, settings.minPrice);
      if (scaledWrapArea > 0) estimate = Math.max(estimate, settings.minAnyWrapPrice);
      if (scaledWrapArea > 0 && scaledCoverage >= 0.15) estimate = Math.max(estimate, settings.minPartialWrapPrice);
      if (scaledWrapArea > 0 && scaledCoverage >= 0.85) estimate = Math.max(estimate, settings.minFullWrapPrice);
      return estimate;
    }

    const currentPrice = calculateScaledEstimate(1);
    const monotonicFloor = Array.from({ length: 80 }, (_, index) => (index + 1) / 80).reduce(
      (highest, scale) => Math.max(highest, calculateScaledEstimate(scale)),
      currentPrice
    );
    return Math.round(Math.max(currentPrice, monotonicFloor) / 50) * 50;
  }

  function getTrainingMetricsFromTotals() {
    return {
      standardPrintArea: totals.standardPrintArea,
      contraArea: totals.contraArea,
      reflectiveArea: totals.reflectiveArea,
      wrapArea: totals.wrapArea,
      totalArea: totals.totalArea,
      vehicleArea: totals.vehicleArea,
      coverage: totals.coverage,
      wrapCoverage: totals.wrapCoverage,
      wrapSectionCount: totals.wrapSectionCount,
      weightedDifficulty: totals.weightedDifficulty
    };
  }

  const totals = useMemo(() => {
    const settings = pricingSettingsOpen ? pricingDraftSettings : pricingSettings;
    const classifiedShapes = shapes.map((shape) => {
      const zoneMetadata =
        shape.zoneMetadata ||
        (shape.isWrapFilm
          ? {
              material_type: "wrap_film",
              surface_type: "normal_wrap_curve",
              complexity_factor: settings.difficultyFactors.normal_wrap_curve,
              install_group: "main_face"
            }
          : {
              material_type: "standard_vinyl",
              surface_type: "flat",
              complexity_factor: settings.difficultyFactors.flat,
              install_group: "standard_panel"
            });
      return { ...shape, materialVariant: getShapeMaterialVariant(shape), zoneMetadata };
    });
    const standardShapes = classifiedShapes.filter((shape) => shape.zoneMetadata.material_type !== "wrap_film");
    const wrapShapes = classifiedShapes.filter((shape) => shape.zoneMetadata.material_type === "wrap_film");
    const standardArea = standardShapes.reduce((sum, shape) => sum + shape.areaM2, 0);
    const standardPrintArea = standardShapes
      .filter((shape) => getShapeMaterialVariant(shape) === "standard")
      .reduce((sum, shape) => sum + shape.areaM2, 0);
    const contraArea = standardShapes
      .filter((shape) => getShapeMaterialVariant(shape) === "contra")
      .reduce((sum, shape) => sum + shape.areaM2, 0);
    const reflectiveArea = standardShapes
      .filter((shape) => getShapeMaterialVariant(shape) === "reflective")
      .reduce((sum, shape) => sum + shape.areaM2, 0);
    const standardMaterialMultiplierArea = standardShapes.reduce((sum, shape) => {
      const multiplier = settings.materialMultipliers[getShapeMaterialVariant(shape)] || settings.materialMultipliers.standard || 1;
      return sum + shape.areaM2 * multiplier;
    }, 0);
    const wrapArea = wrapShapes.reduce((sum, shape) => sum + shape.areaM2, 0);
    const totalArea = standardArea + wrapArea;
    const vehicleArea = getVehicleDrawableAreaM2();
    const coverage = vehicleArea > 0 ? totalArea / vehicleArea : 0;
    const wrapCoverage = vehicleArea > 0 ? wrapArea / vehicleArea : 0;
    const wrapSectionCount = countConnectedShapeSections(wrapShapes);
    const wrapSectionsConnected = wrapSectionCount <= 1;
    const wrapSectionFactor = wrapArea > 0 ? getWrapSectionFactor(wrapSectionCount, settings) : 1;
    const weightedDifficulty = getWeightedDifficultyFactor(wrapShapes, wrapArea, settings);

    if (totalArea <= 0) {
      return {
        standardArea,
        standardPrintArea,
        contraArea,
        reflectiveArea,
        wrapArea,
        totalArea,
        standardSell: 0,
        standardLabourSell: 0,
        wrapSell: 0,
        wrapLabourSell: 0,
        labourHours: 0,
        basePrice: 0,
        vehicleArea,
        coverage: 0,
        wrapCoverage: 0,
        wrapSectionCount: 0,
        wrapSectionsConnected: true,
        weightedDifficulty: 1,
        anchor: 0,
        packageType: "",
        estimate: 0
      };
    }
    function calculateScaledEstimate(scale = 1) {
      const scaledStandardArea = standardArea * scale;
      const scaledStandardMaterialMultiplierArea = standardMaterialMultiplierArea * scale;
      const scaledWrapArea = wrapArea * scale;
      const scaledTotalArea = totalArea * scale;
      const scaledCoverage = vehicleArea > 0 ? scaledTotalArea / vehicleArea : 0;
      const scaledWrapCoverage = vehicleArea > 0 ? scaledWrapArea / vehicleArea : 0;
      const standardSell = scaledStandardMaterialMultiplierArea * settings.standardVinylRate;
      const wrapFilmRate = getFormulaWrapRate(scaledWrapCoverage, settings);
      const wrapSell = scaledWrapArea * wrapFilmRate;
      const standardLabourHours =
        scaledStandardArea <= 0
          ? 0
          : scaledCoverage < 0.15
            ? clampNumber(
                scaledStandardArea * settings.standardSmallHoursPerM2,
                settings.standardSmallMinHours,
                settings.standardSmallMaxHours
              )
            : scaledStandardArea * settings.standardLargeHoursPerM2;
      const wrapBaseHoursPerM2 = getFormulaWrapLabourHours(scaledWrapCoverage, settings);
      const wrapLabourHours = scaledWrapArea * wrapBaseHoursPerM2 * wrapSectionFactor * weightedDifficulty;
      const labourHours = standardLabourHours + wrapLabourHours;
      const labourSell = labourHours * settings.labourSellRate;
      const standardLabourSell = standardLabourHours * settings.labourSellRate;
      const wrapLabourSell = wrapLabourHours * settings.labourSellRate;
      const basePrice = standardSell + standardLabourSell + wrapSell + wrapLabourSell;
      const anchor = getMarketAnchor(scaledCoverage, settings);
      const blend = getBlendWeights(scaledCoverage, scaledWrapArea, settings);
      let estimate = blend.calculated * basePrice + blend.anchor * anchor;
      estimate = Math.max(estimate, settings.minPrice);
      if (scaledWrapArea > 0) {
        estimate = Math.max(estimate, settings.minAnyWrapPrice);
      }
      if (scaledWrapArea > 0 && scaledCoverage >= 0.15) {
        estimate = Math.max(estimate, settings.minPartialWrapPrice);
      }
      if (scaledWrapArea > 0 && scaledCoverage >= 0.85) {
        estimate = Math.max(estimate, settings.minFullWrapPrice);
      }

      return {
        anchor,
        basePrice,
        calculatedPrice: basePrice,
        estimate,
        labourHours,
        labourSell,
        packageType: "",
        standardLabourSell,
        standardSell,
        wrapFilmRate,
        wrapBaseHoursPerM2,
        wrapLabourSell,
        wrapSell
      };
    }

    const currentPrice = calculateScaledEstimate(1);
    const monotonicFloor = Array.from({ length: 80 }, (_, index) => (index + 1) / 80).reduce((highest, scale) => {
      return Math.max(highest, calculateScaledEstimate(scale).estimate);
    }, currentPrice.estimate);
    const monotonicEstimate = Math.max(currentPrice.estimate, monotonicFloor);
    const estimate = Math.round(monotonicEstimate / 50) * 50;

    return {
      standardArea,
      standardPrintArea,
      contraArea,
      reflectiveArea,
      wrapArea,
      totalArea,
      standardSell: currentPrice.standardSell,
      standardLabourSell: currentPrice.standardLabourSell,
      wrapSell: currentPrice.wrapSell,
      wrapLabourSell: currentPrice.wrapLabourSell,
      labourHours: currentPrice.labourHours,
      basePrice: currentPrice.basePrice,
      vehicleArea,
      coverage,
      wrapCoverage,
      wrapSectionCount,
      wrapSectionsConnected,
      weightedDifficulty,
      anchor: currentPrice.anchor,
      packageType: currentPrice.packageType,
      estimate
    };
  }, [activeScaleFactor, pricingDraftSettings, pricingSettings, pricingSettingsOpen, selectedTemplate.id, shapes, vehicleClipPathsD.length]);

  const currencyFormatter = useMemo(
    () =>
      new Intl.NumberFormat("en-GB", {
        style: "currency",
        currency: "GBP",
        maximumFractionDigits: 0
      }),
    []
  );

  function formatM2(value) {
    return `${(Number(value) || 0).toFixed(2)}m²`;
  }

  function formatPercent(value) {
    return `${Math.round((Number(value) || 0) * 100)}%`;
  }

  function getShapeVisualClass(shape) {
    if (shape.isWrapFilm) return "wrap";
    return getShapeMaterialVariant(shape);
  }

  function getShapeCenter(shape) {
    const bounds = shape.bounds || shape;
    return {
      x: bounds.x + bounds.width / 2,
      y: bounds.y + bounds.height / 2
    };
  }

  function getExportShapeColours(shape) {
    const visualClass = getShapeVisualClass(shape);
    if (visualClass === "wrap") return { fill: "#fdba74", opacity: 0.34, stroke: "#ea580c" };
    if (visualClass === "contra") return { fill: "#0f172a", opacity: 0.18, stroke: "#0f172a" };
    if (visualClass === "reflective") return { fill: "#ffffff", opacity: 0.74, stroke: "#64748b" };
    return { fill: "#7dd3fc", opacity: 0.24, stroke: "#0284c7" };
  }

  function wrapSvgText(value, maxLength = 24, maxLines = 10) {
    const words = String(value || "-").replace(/\s+/g, " ").trim().split(" ").filter(Boolean);
    const lines = [];
    let currentLine = "";

    words.forEach((word) => {
      const nextLine = currentLine ? `${currentLine} ${word}` : word;
      if (nextLine.length <= maxLength) {
        currentLine = nextLine;
        return;
      }
      if (currentLine) lines.push(currentLine);
      currentLine = word;
    });

    if (currentLine) lines.push(currentLine);
    const trimmedLines = lines.slice(0, maxLines);
    if (lines.length > maxLines && trimmedLines.length) {
      trimmedLines[trimmedLines.length - 1] = `${trimmedLines[trimmedLines.length - 1].slice(0, Math.max(0, maxLength - 3))}...`;
    }
    return trimmedLines.length ? trimmedLines : ["-"];
  }

  function createExportTextBlock({ x, y, title, value, maxLength = 24, maxLines = 3, anchor = "start", valueSize = 8.2, titleSize = 8.5, valueWeight = 400 }) {
    const valueLines = wrapSvgText(value, maxLength, maxLines);
    const titleMarkup = `<text x="${x}" y="${y}" text-anchor="${anchor}" fill="#5f3c74" font-family="Faricy, 'Faricy New', Arial, sans-serif" font-size="${titleSize}" font-weight="700">${escapeSvgText(title)}</text>`;
    const valueMarkup = valueLines
      .map(
        (line, index) =>
          `<text x="${x}" y="${y + 11 + index * 9}" text-anchor="${anchor}" fill="#172033" font-family="Faricy, 'Faricy New', Arial, sans-serif" font-size="${valueSize}" font-weight="${valueWeight}">${escapeSvgText(line)}</text>`
      )
      .join("");
    return { markup: `${titleMarkup}${valueMarkup}`, height: 14 + valueLines.length * 9 };
  }

  function buildExportShapeMarkup() {
    return shapes
      .map((shape) => {
        const colours = getExportShapeColours(shape);
        const common = `fill="${colours.fill}" fill-opacity="${colours.opacity}" stroke="${colours.stroke}" stroke-width="4" vector-effect="non-scaling-stroke" stroke-linejoin="round" stroke-linecap="round"`;
        if (shape.type === "polygon") {
          return `<polygon points="${pointsToSvg(shape.points)}" ${common} />`;
        }
        return `<rect x="${shape.x}" y="${shape.y}" width="${shape.width}" height="${shape.height}" rx="10" ${common} />`;
      })
      .join("");
  }

  function svgToDataUri(svg) {
    return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg.trim())}`;
  }

  function buildVehicleExportSvg() {
    const sourceSvg = inlineSvgRef.current?.querySelector("svg");
    if (!sourceSvg) return "";
    const viewBox = selectedTemplate.viewBox;
    const clipMarkup = vehicleClipPathsD.length
      ? `<defs><clipPath id="export-vehicle-body-clip">${vehicleClipPathsD
          .map((pathD) => `<path d="${pathD}" clip-rule="evenodd" />`)
          .join("")}</clipPath></defs>`
      : "";
    const clipAttribute = vehicleClipPathsD.length ? ` clip-path="url(#export-vehicle-body-clip)"` : "";
    return `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}">
        ${clipMarkup}
        <svg x="${viewBox.x}" y="${viewBox.y}" width="${viewBox.width}" height="${viewBox.height}" viewBox="${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}" preserveAspectRatio="xMidYMid meet">
          ${sourceSvg.innerHTML}
        </svg>
        <g id="export-drawn-shapes"${clipAttribute}>${buildExportShapeMarkup()}</g>
      </svg>
    `;
  }

  function buildArtBoardExportSvg(customerName, notes) {
    if (!artBoardMarkup) throw new Error("The art board SVG has not loaded yet.");
    const vehicleSvg = buildVehicleExportSvg();
    if (!vehicleSvg) throw new Error("The vehicle artwork has not loaded yet.");

    const parser = new DOMParser();
    const artBoardDocument = parser.parseFromString(artBoardMarkup, "image/svg+xml");
    const parseError = artBoardDocument.querySelector("parsererror");
    if (parseError) throw new Error("Could not read the art board SVG.");

    const artBoardSvg = artBoardDocument.querySelector("svg");
    const vehicleArea = artBoardDocument.querySelector("#Area_to_put_van rect");
    if (!artBoardSvg || !vehicleArea) throw new Error("The art board is missing the van placement area.");

    const area = {
      x: Number(vehicleArea.getAttribute("x") || 0),
      y: Number(vehicleArea.getAttribute("y") || 0),
      width: Number(vehicleArea.getAttribute("width") || 0),
      height: Number(vehicleArea.getAttribute("height") || 0)
    };
    const viewBox = selectedTemplate.viewBox;
    const scale = Math.min(area.width / viewBox.width, area.height / viewBox.height);
    const width = viewBox.width * scale;
    const height = viewBox.height * scale;
    const x = area.x + (area.width - width) / 2;
    const y = area.y + (area.height - height) / 2;
    const vehicleImage = artBoardDocument.createElementNS("http://www.w3.org/2000/svg", "image");
    vehicleImage.setAttribute("id", "Exported_Van");
    vehicleImage.setAttribute("x", String(x));
    vehicleImage.setAttribute("y", String(y));
    vehicleImage.setAttribute("width", String(width));
    vehicleImage.setAttribute("height", String(height));
    vehicleImage.setAttribute("preserveAspectRatio", "xMidYMid meet");
    vehicleImage.setAttribute("href", svgToDataUri(vehicleSvg));
    vehicleImage.setAttributeNS("http://www.w3.org/1999/xlink", "xlink:href", svgToDataUri(vehicleSvg));

    const detailBox = artBoardDocument.querySelector("#Detail_Box");
    artBoardSvg.insertBefore(vehicleImage, detailBox || null);

    artBoardDocument.querySelector("#Draft_x2C__Date_x2C__Designer")?.remove();

    const exportDate = new Intl.DateTimeFormat("en-GB", {
      day: "2-digit",
      month: "2-digit",
      year: "2-digit"
    }).format(new Date());
    const designerName = String(currentUser?.displayName || currentUser?.name || currentUser?.username || "Matt").trim().split(/\s+/)[0] || "Matt";
    Array.from(artBoardDocument.querySelectorAll("text")).forEach((textNode) => {
      const label = String(textNode.textContent || "").trim().toLowerCase();
      if (["draft", "date", "designer"].includes(label)) {
        textNode.setAttribute("font-family", "Faricy, 'Faricy New', Arial, sans-serif");
        textNode.setAttribute("font-weight", "700");
      }
    });
    const metaGroup = artBoardDocument.createElementNS("http://www.w3.org/2000/svg", "g");
    metaGroup.setAttribute("id", "Draft_x2C__Date_x2C__Designer");
    metaGroup.innerHTML = `
      <text x="721.99" y="556.64" text-anchor="middle" fill="#fff" font-family="Faricy, 'Faricy New', Arial, sans-serif" font-size="8.5" font-weight="700">1</text>
      <text x="764.28" y="556.64" text-anchor="middle" fill="#fff" font-family="Faricy, 'Faricy New', Arial, sans-serif" font-size="8.5" font-weight="700">${escapeSvgText(exportDate)}</text>
      <text x="806.57" y="556.64" text-anchor="middle" fill="#fff" font-family="Faricy, 'Faricy New', Arial, sans-serif" font-size="8.5" font-weight="700">${escapeSvgText(designerName)}</text>
    `;
    artBoardSvg.appendChild(metaGroup);

    const priceLabel = `${currencyFormatter.format(totals.estimate || 0)} +VAT`;
    const details = [
      ["Customer", customerName || "-"],
      ["Std. print vinyl", formatM2(totals.standardPrintArea)],
      ["Wrap film", formatM2(totals.wrapArea)],
      ["Contra-vision", formatM2(totals.contraArea)],
      ["Reflective", formatM2(totals.reflectiveArea)],
      ["Total coverage", formatM2(totals.totalArea)],
      ["% coverage", formatPercent(totals.coverage)],
      ["Est. application", `${(Number(totals.labourHours) || 0).toFixed(1)} hours`]
    ];
    let detailY = 108;
    const priceBlock = createExportTextBlock({
      x: 822,
      y: detailY,
      title: "Price",
      value: priceLabel,
      maxLength: 18,
      maxLines: 1,
      anchor: "end",
      valueSize: 13,
      titleSize: 9,
      valueWeight: 700
    });
    detailY += priceBlock.height + 8;
    const detailMarkup = details
      .map(([title, value]) => {
        const block = createExportTextBlock({ x: 822, y: detailY, title, value, maxLength: 22, maxLines: 2, anchor: "end" });
        detailY += block.height + 2;
        return block.markup;
      })
      .join("");
    const notesBlock = createExportTextBlock({
      x: 822,
      y: detailY + 5,
      title: "Notes",
      value: notes || "-",
      maxLength: 26,
      maxLines: 12,
      anchor: "end"
    });
    const detailTextGroup = artBoardDocument.createElementNS("http://www.w3.org/2000/svg", "g");
    detailTextGroup.setAttribute("id", "Export_Detail_Text");
    detailTextGroup.innerHTML = `${priceBlock.markup}${detailMarkup}${notesBlock.markup}`;
    artBoardSvg.appendChild(detailTextGroup);

    return new XMLSerializer().serializeToString(artBoardSvg);
  }

  function openPdfExport() {
    setPdfExportForm({ customerName: "", notes: "" });
    setPdfExportError("");
    setPdfExportOpen(true);
  }

  async function exportVehiclePdf(event) {
    event.preventDefault();
    const printWindow = window.open("", "_blank");
    if (!printWindow) {
      setPdfExportError("Please allow pop-ups so the PDF export can open.");
      return;
    }

    setPdfExporting(true);
    setPdfExportError("");
    setSvgError("");
    try {
      printWindow.document.write(`<!doctype html><html><head><title>Creating PDF...</title></head><body style="font-family:Arial,sans-serif;padding:24px;">Creating PDF...</body></html>`);
      printWindow.document.close();
      const exportSvg = buildArtBoardExportSvg(pdfExportForm.customerName, pdfExportForm.notes);
      printWindow.document.open();
      printWindow.document.write(`<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(pdfExportForm.customerName || "Vehicle Proposal")}</title>
    <style>
      @font-face {
        font-family: 'Faricy';
        src: url('${window.location.origin}/fonts/FARICYNEW-MEDIUM.TTF') format('truetype');
        font-weight: 400;
      }
      @font-face {
        font-family: 'Faricy';
        src: url('${window.location.origin}/fonts/FARICYNEW-BOLD.TTF') format('truetype');
        font-weight: 700;
      }
      @page { size: A4 landscape; margin: 0; }
      html, body {
        margin: 0;
        width: 100%;
        min-height: 100%;
        background: #fff;
        font-family: Faricy, Arial, sans-serif;
        print-color-adjust: exact;
        -webkit-print-color-adjust: exact;
      }
      .sheet { width: 297mm; height: 210mm; margin: 0 auto; overflow: hidden; }
      .sheet svg {
        display: block;
        width: 297mm;
        height: 210mm;
        print-color-adjust: exact;
        -webkit-print-color-adjust: exact;
      }
    </style>
  </head>
  <body>
    <div class="sheet">${exportSvg}</div>
    <script>
      window.addEventListener('load', () => {
        setTimeout(() => {
          window.focus();
          window.print();
        }, 250);
      });
    </script>
  </body>
</html>`);
      printWindow.document.close();
      setPdfExportOpen(false);
    } catch (error) {
      console.error(error);
      printWindow.close();
      setPdfExportError(error.message || "Could not export the vehicle PDF.");
    } finally {
      setPdfExporting(false);
    }
  }

  function updatePricingValue(path, value) {
    const numericValue = Number(value);
    if (Number.isNaN(numericValue)) return;
    setPricingDraftSettings((current) => setNestedPricingValue(current, path, numericValue));
  }

  function renderPricingNumber(label, path, step = 1, help = "") {
    const value = path.reduce((target, key) => target?.[key], pricingDraftSettings);
    return (
      <label>
        <span className={help ? "pricing-label-help" : ""} data-help={help || undefined}>
          {label}
        </span>
        <input type="number" step={step} value={value} onChange={(event) => updatePricingValue(path, event.target.value)} />
      </label>
    );
  }

  function importSmartPrices() {
    const result = parseSmartPriceImport(smartPriceImportText, selectedTemplate);
    if (!result.updates.length) {
      setSmartPriceImportStatus(`No matching ${selectedTemplate.sizeName} pricing lines found.`);
      return;
    }

    setPricingDraftSettings((current) =>
      result.updates.reduce((nextSettings, update) => setNestedPricingValue(nextSettings, update.path, update.value), current)
    );
    setSmartPriceImportStatus(
      `Imported ${result.updates.length} ${selectedTemplate.sizeName} value${result.updates.length === 1 ? "" : "s"}. Press Save to keep them.`
    );
  }

  function addTrainingExample() {
    const targetPrice = Number(trainingTargetPrice);
    if (!shapes.length || totals.totalArea <= 0) {
      setTrainingStatus("Draw at least one priced area first.");
      return;
    }
    if (!Number.isFinite(targetPrice) || targetPrice <= 0) {
      setTrainingStatus("Enter the manual price you would charge for this layout.");
      return;
    }

    const example = {
      id: `training-${Date.now()}-${Math.round(targetPrice)}`,
      templateId: selectedTemplate.id,
      templateName: selectedTemplate.sizeName,
      targetPrice,
      currentEstimate: totals.estimate,
      metrics: getTrainingMetricsFromTotals(),
      createdAt: new Date().toISOString()
    };

    setPriceTrainingBank((current) => [example, ...current]);
    setTrainingTargetPrice("");
    setTrainingSuggestion(null);
    setTrainingStatus(`Added ${currencyFormatter.format(targetPrice)} example to the ${selectedTemplate.sizeName} bank.`);
  }

  function removeTrainingExample(exampleId) {
    setPriceTrainingBank((current) => current.filter((example) => example.id !== exampleId));
    setTrainingSuggestion(null);
  }

  function getWeightedTrainingRatio(samples, sourceSettings) {
    let weightedRatio = 0;
    let totalWeight = 0;

    samples.forEach((sample) => {
      const estimate = calculateEstimateFromMetrics(sample.metrics, sourceSettings);
      const target = Number(sample.targetPrice) || 0;
      if (!estimate || !target) return;
      const weight = Math.max(0.25, Number(sample.metrics?.totalArea) || 0.25);
      weightedRatio += (target / estimate) * weight;
      totalWeight += weight;
    });

    return totalWeight > 0 ? weightedRatio / totalWeight : 1;
  }

  function getSettingsFitScore(samples, candidateSettings) {
    if (!samples.length) return 0;
    const totalError = samples.reduce((sum, sample) => {
      const estimate = calculateEstimateFromMetrics(sample.metrics, candidateSettings);
      const target = Number(sample.targetPrice) || 0;
      if (!target) return sum;
      return sum + Math.abs(estimate - target) / target;
    }, 0);
    return totalError / samples.length;
  }

  function scaleSettingGroup(settings, keys, ratio) {
    return keys.reduce((nextSettings, path) => {
      const currentValue = path.reduce((target, key) => target?.[key], nextSettings);
      if (!Number.isFinite(Number(currentValue))) return nextSettings;
      return setNestedPricingValue(nextSettings, path, Math.round(Number(currentValue) * ratio * 100) / 100);
    }, settings);
  }

  function calculateTrainingSuggestion() {
    const samples = selectedTemplateTrainingBank;
    if (!samples.length) {
      setTrainingStatus(`Add at least one ${selectedTemplate.sizeName} example first.`);
      return;
    }

    const ratio = clampNumber(getWeightedTrainingRatio(samples, pricingDraftSettings), 0.5, 1.8);
    const anchorPaths = [
      ["marketAnchors", "c05"],
      ["marketAnchors", "c10"],
      ["marketAnchors", "c15"],
      ["marketAnchors", "c22"],
      ["marketAnchors", "c35"],
      ["marketAnchors", "c55"],
      ["marketAnchors", "c85"],
      ["marketAnchors", "c100"]
    ];
    const ratePaths = [
      ["standardVinylRate"],
      ["wrapRateStart"],
      ["wrapRateFloor"],
      ["labourSellRate"],
      ["minPrice"],
      ["minAnyWrapPrice"],
      ["minPartialWrapPrice"],
      ["minFullWrapPrice"]
    ];
    const balancedRatio = ratio;
    const candidates = [
      {
        label: "Balanced",
        settings: scaleSettingGroup(scaleSettingGroup(pricingDraftSettings, anchorPaths, balancedRatio), ratePaths, balancedRatio)
      },
      { label: "Anchors only", settings: scaleSettingGroup(pricingDraftSettings, anchorPaths, ratio) },
      { label: "Rates & minimums", settings: scaleSettingGroup(pricingDraftSettings, ratePaths, ratio) }
    ].map((candidate) => ({
      ...candidate,
      score: getSettingsFitScore(samples, candidate.settings)
    }));
    const bestCandidate = candidates.sort((left, right) => left.score - right.score)[0];
    const beforeScore = getSettingsFitScore(samples, pricingDraftSettings);

    setTrainingSuggestion({
      ...bestCandidate,
      beforeScore,
      ratio,
      samples: samples.length
    });
    setTrainingStatus(
      `${bestCandidate.label} suggestion ready from ${samples.length} example${samples.length === 1 ? "" : "s"}. Review it, then apply if it looks right.`
    );
  }

  function applyTrainingSuggestion() {
    if (!trainingSuggestion?.settings) return;
    setPricingDraftSettings(mergeVehiclePricingSettings(trainingSuggestion.settings));
    setTrainingStatus("Suggested values applied to the draft pricing settings. Press Save to keep them.");
  }

  async function savePricingSettings() {
    const nextSettings = mergeVehiclePricingSettings(pricingDraftSettings);
    setPricingDraftSettings(nextSettings);
    const nextSettingsByTemplate = {
      ...pricingSettingsByTemplate,
      [selectedTemplate.id]: nextSettings
    };
    setPricingSettingsByTemplate((current) => {
      const nextSettingsByTemplateForState = {
        ...current,
        [selectedTemplate.id]: nextSettings
      };
      saveVehiclePricingSettingsByTemplate(nextSettingsByTemplateForState);
      return nextSettingsByTemplateForState;
    });
    setPricingSyncStatus("Saving shared pricing...");
    try {
      const response = await fetch("/api/vehicle-pricing/settings", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ settingsByTemplate: nextSettingsByTemplate })
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.error || "Could not save shared vehicle pricing.");
      const sharedSettings = normalizeVehiclePricingSettingsByTemplate(payload.settingsByTemplate || nextSettingsByTemplate);
      setPricingSettingsByTemplate(sharedSettings);
      saveVehiclePricingSettingsByTemplate(sharedSettings);
      setPricingSyncStatus("Shared pricing saved.");
    } catch (error) {
      console.error(error);
      setPricingSyncStatus("Saved in this browser, but shared pricing did not update.");
    }
  }

  function resetPricingDraft() {
    setPricingDraftSettings(pricingSettings);
  }

  function handlePricingSettingsToggle(event) {
    const isOpen = event.currentTarget.open;
    setPricingSettingsOpen(isOpen);
    setPricingDraftSettings(pricingSettings);
  }

  async function restoreDefaultPricingSettings() {
    const defaultSettings = mergeVehiclePricingSettings(selectedTemplate.pricingSettings || {});
    setPricingDraftSettings(defaultSettings);
    const nextSettingsByTemplate = {
      ...pricingSettingsByTemplate,
      [selectedTemplate.id]: defaultSettings
    };
    setPricingSettingsByTemplate((current) => {
      const nextSettingsByTemplateForState = {
        ...current,
        [selectedTemplate.id]: defaultSettings
      };
      saveVehiclePricingSettingsByTemplate(nextSettingsByTemplateForState);
      return nextSettingsByTemplateForState;
    });
    setPricingSyncStatus("Saving shared defaults...");
    try {
      const response = await fetch("/api/vehicle-pricing/settings", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ settingsByTemplate: nextSettingsByTemplate })
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.error || "Could not save shared vehicle pricing.");
      const sharedSettings = normalizeVehiclePricingSettingsByTemplate(payload.settingsByTemplate || nextSettingsByTemplate);
      setPricingSettingsByTemplate(sharedSettings);
      saveVehiclePricingSettingsByTemplate(sharedSettings);
      setPricingSyncStatus("Shared defaults saved.");
    } catch (error) {
      console.error(error);
      setPricingSyncStatus("Defaults saved in this browser, but shared pricing did not update.");
    }
  }

  return (
    <div className="app-shell">
      <div className="page vinyl-estimator-page">
        <MainNavBar
          currentUser={currentUser}
          active="van-estimator"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel vinyl-estimator-panel">
          <div className="vinyl-estimator-grid">
            <div className="vinyl-canvas-card">
              {svgError ? <div className="flash error">{svgError}</div> : null}
              <div className="vinyl-canvas">
                <button
                  className="vinyl-pdf-button"
                  type="button"
                  onClick={openPdfExport}
                  onPointerDown={(event) => {
                    event.stopPropagation();
                  }}
                >
                  PDF
                </button>
                <div
                  className="vinyl-canvas-toolbar"
                  aria-label="Drawing tools"
                  onPointerDown={(event) => {
                    event.stopPropagation();
                  }}
                >
                  <button
                    type="button"
                    className={drawMode === "rectangle" ? "active" : ""}
                    title="Rectangle tool"
                    aria-label="Rectangle tool"
                    onClick={() => selectDrawMode("rectangle")}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <rect x="5" y="6" width="14" height="12" rx="1.5" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className={drawMode === "polygon" ? "active" : ""}
                    title="Point shape tool"
                    aria-label="Point shape tool"
                    onClick={() => selectDrawMode("polygon")}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <path d="M6 17 9 6l9 4-2 8-10-1Z" />
                      <circle cx="9" cy="6" r="1.7" />
                      <circle cx="18" cy="10" r="1.7" />
                      <circle cx="16" cy="18" r="1.7" />
                      <circle cx="6" cy="17" r="1.7" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className={drawMode === "lasso" ? "active" : ""}
                    title="Draw shape tool"
                    aria-label="Draw shape tool"
                    onClick={() => selectDrawMode("lasso")}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <path d="M6.5 13.2c0-3 2.7-5.4 6-5.4s6 2.4 6 5.4-2.7 5.4-6 5.4-6-2.4-6-5.4Z" />
                      <path d="M12.5 7.8c1.4-2.4 3.3-3.4 5.8-3" />
                      <path d="M18.3 4.8 17 3.6" />
                      <path d="M18.3 4.8 16.6 5.7" />
                      <circle cx="12.5" cy="13.2" r="1.3" />
                    </svg>
                  </button>
                  <span className="vinyl-toolbar-divider" aria-hidden="true" />
                  <button
                    type="button"
                    title="Undo last"
                    aria-label="Undo last"
                    disabled={!shapes.length && !polygonPoints.length && !lassoPoints.length}
                    onClick={undoLastDrawing}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <path d="M9 7 5 11l4 4" />
                      <path d="M5 11h8.2c3.2 0 5.8 2.4 5.8 5.4 0 1.1-.3 2.1-.9 2.9" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    title="Clear"
                    aria-label="Clear all"
                    disabled={!shapes.length && !polygonPoints.length && !lassoPoints.length}
                    onClick={clearDrawing}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <path d="m7 7 10 10" />
                      <path d="m17 7-10 10" />
                    </svg>
                  </button>
                </div>
                <div
                  className="vinyl-material-toolbar"
                  aria-label="Material modifiers"
                  onPointerDown={(event) => {
                    event.stopPropagation();
                  }}
                >
                  <button
                    type="button"
                    className={materialMode === "standard" ? "active" : ""}
                    title="Standard vinyl"
                    aria-label="Standard vinyl"
                    onClick={() => setMaterialMode("standard")}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <rect x="5" y="5" width="14" height="14" rx="2" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className={materialMode === "contra" ? "active" : ""}
                    title="Contra-vision"
                    aria-label="Contra-vision"
                    onClick={() => setMaterialMode("contra")}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <rect x="5" y="5" width="14" height="14" rx="2" />
                      <circle cx="9" cy="9" r="1" />
                      <circle cx="15" cy="9" r="1" />
                      <circle cx="12" cy="12" r="1" />
                      <circle cx="9" cy="15" r="1" />
                      <circle cx="15" cy="15" r="1" />
                    </svg>
                  </button>
                  <button
                    type="button"
                    className={materialMode === "reflective" ? "active" : ""}
                    title="Reflective"
                    aria-label="Reflective"
                    onClick={() => setMaterialMode("reflective")}
                  >
                    <svg viewBox="0 0 24 24" aria-hidden="true">
                      <rect x="5" y="5" width="14" height="14" rx="2" />
                      <path d="m9 15 6-6" />
                      <path d="m12 16 4-4" />
                    </svg>
                  </button>
                </div>
                <div
                  ref={inlineSvgRef}
                  className="vinyl-template"
                  dangerouslySetInnerHTML={{ __html: svgMarkup }}
                />
                <svg
                  ref={overlaySvgRef}
                  className="vinyl-drawing-layer"
                  viewBox={`${selectedTemplate.viewBox.x} ${selectedTemplate.viewBox.y} ${selectedTemplate.viewBox.width} ${selectedTemplate.viewBox.height}`}
                  preserveAspectRatio="xMidYMid meet"
                  onPointerDown={startDrawing}
                  onPointerMove={updateDrawing}
                  onPointerUp={finishDrawing}
                  onPointerCancel={() => {
                    setDrawStart(null);
                    setDrawingRect(null);
                    setEditDrag(null);
                    setLassoPoints([]);
                  }}
                >
                  {vehicleClipPathsD.length ? (
                    <defs>
                      <pattern id="contra-vision-dot-pattern" width="12" height="12" patternUnits="userSpaceOnUse">
                        <rect width="12" height="12" fill="rgba(15, 23, 42, 0.72)" />
                        <circle cx="3" cy="3" r="1.4" fill="rgba(255, 255, 255, 0.78)" />
                        <circle cx="9" cy="9" r="1.4" fill="rgba(255, 255, 255, 0.78)" />
                      </pattern>
                      <clipPath id="vinyl-vehicle-body-clip">
                        {vehicleClipPathsD.map((pathD, index) => (
                          <path key={`vehicle-body-clip-${index}`} d={pathD} clipRule="evenodd" />
                        ))}
                      </clipPath>
                      {shapes.map((shape) => (
                        <clipPath key={`shape-clip-${shape.id}`} id={getShapeClipId(shape.id)}>
                          {shape.type === "polygon" ? (
                            <polygon points={pointsToSvg(shape.points)} />
                          ) : (
                            <rect x={shape.x} y={shape.y} width={shape.width} height={shape.height} />
                          )}
                        </clipPath>
                      ))}
                    </defs>
                  ) : null}
                  {shapes.map((shape) => {
                    const shapeCenter = getShapeCenter(shape);
                    return (
                      <g key={shape.id} className="vinyl-shape-group">
                        <g clipPath={vehicleClipPathsD.length ? "url(#vinyl-vehicle-body-clip)" : undefined}>
                          {shape.type === "polygon" ? (
                            <polygon
                              points={pointsToSvg(shape.points)}
                              className={`vinyl-shape ${getShapeVisualClass(shape)}`}
                            />
                          ) : (
                            <rect
                              x={shape.x}
                              y={shape.y}
                              width={shape.width}
                              height={shape.height}
                              className={`vinyl-shape ${getShapeVisualClass(shape)}`}
                            />
                          )}
                        </g>
                        {vehicleEdgePathsD.length ? (
                          <g
                            className={`vinyl-vehicle-edge ${getShapeVisualClass(shape)}`}
                            clipPath={`url(#${getShapeClipId(shape.id)})`}
                          >
                            {vehicleEdgePathsD.map((pathD, index) => (
                              <path key={`${shape.id}-edge-${index}`} d={pathD} />
                            ))}
                          </g>
                        ) : null}
                        {shape.type === "polygon" ? (
                          <polygon
                            points={pointsToSvg(shape.points)}
                            className={`vinyl-shape-cutline ${getShapeVisualClass(shape)}`}
                            clipPath={vehicleClipPathsD.length ? "url(#vinyl-vehicle-body-clip)" : undefined}
                          />
                        ) : (
                          <rect
                            x={shape.x}
                            y={shape.y}
                            width={shape.width}
                            height={shape.height}
                            className={`vinyl-shape-cutline ${getShapeVisualClass(shape)}`}
                            clipPath={vehicleClipPathsD.length ? "url(#vinyl-vehicle-body-clip)" : undefined}
                          />
                        )}
                        <g className="vinyl-shape-controls" transform={`translate(${shapeCenter.x} ${shapeCenter.y})`}>
                          <g
                            className={`vinyl-shape-control vinyl-shape-material-toggle ${getShapeVisualClass(shape)}`}
                            role="button"
                            tabIndex="0"
                            aria-label={shape.isWrapFilm ? "Change drawn area to standard vinyl" : "Change drawn area to wrap film"}
                            onPointerDown={(event) => {
                              event.stopPropagation();
                              event.preventDefault();
                            }}
                            onClick={(event) => toggleShapeMaterial(event, shape.id)}
                          >
                            <circle cx="-22" cy="0" r="15" />
                            <line x1="-30" y1="0" x2="-14" y2="0" />
                            <polyline points="-25,-6 -31,0 -25,6" />
                            <polyline points="-19,-6 -13,0 -19,6" />
                          </g>
                          <g
                            className="vinyl-shape-control vinyl-shape-delete"
                            role="button"
                            tabIndex="0"
                            aria-label="Delete drawn area"
                            onPointerDown={(event) => {
                              event.stopPropagation();
                              event.preventDefault();
                            }}
                            onClick={(event) => deleteShape(event, shape.id)}
                          >
                            <circle cx="22" cy="0" r="15" />
                            <line x1="16" y1="-6" x2="28" y2="6" />
                            <line x1="28" y1="-6" x2="16" y2="6" />
                          </g>
                        </g>
                        {(shape.type === "polygon" ? shape.points : getRectanglePoints(shape.bounds || shape)).map((point, pointIndex) => (
                          <circle
                            key={`${shape.id}-handle-${pointIndex}`}
                            cx={point.x}
                            cy={point.y}
                            r="8"
                            className="vinyl-corner-handle"
                            onPointerDown={(event) =>
                              startShapeCornerDrag(event, shape, shape.type === "polygon" ? pointIndex : point.corner)
                            }
                          />
                        ))}
                      </g>
                    );
                  })}
                  {drawingRect ? (
                    <rect
                      x={drawingRect.x}
                      y={drawingRect.y}
                      width={drawingRect.width}
                      height={drawingRect.height}
                      className={`vinyl-shape drawing ${materialMode}`}
                    />
                  ) : null}
                  {polygonPoints.length ? (
                    <g>
                      <polyline
                        points={pointsToSvg(polygonPreviewPoint ? [...polygonPoints, polygonPreviewPoint] : polygonPoints)}
                        className={`vinyl-shape drawing vinyl-polygon-preview ${materialMode}`}
                      />
                      {polygonPoints.map((point, index) => (
                        <circle
                          key={`polygon-point-${index}`}
                          cx={point.x}
                          cy={point.y}
                          r={index === 0 && polygonPoints.length >= 3 ? "12" : "7"}
                          className={`vinyl-point-handle ${index === 0 && polygonPoints.length >= 3 ? "snap-target" : ""}`}
                        />
                      ))}
                    </g>
                  ) : null}
                  {lassoPoints.length ? (
                    <polyline points={pointsToSvg(lassoPoints)} className={`vinyl-shape drawing vinyl-lasso-preview ${materialMode}`} />
                  ) : null}
                </svg>
              </div>
            </div>

            <aside className="vinyl-estimate-card">
              <label className="vinyl-estimator-template">
                <span>Select Vehicle Size:</span>
                <select value={selectedTemplateId} onChange={(event) => setSelectedTemplateId(event.target.value)}>
                  {VEHICLE_TEMPLATE_OPTIONS.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.sizeName}
                    </option>
                  ))}
                </select>
                <small>eg. {selectedTemplate.exampleName}</small>
              </label>

              <div className="vinyl-total-hero">
                <span>Estimated supply & install ex VAT</span>
                <strong>{currencyFormatter.format(totals.estimate)}</strong>
              </div>

              <div className="vinyl-stats-grid">
                <div className="standard">
                  <span>Std. Print Vinyl</span>
                  <strong>{formatM2(totals.standardPrintArea)}</strong>
                </div>
                <div className="wrap">
                  <span>Wrap film</span>
                  <strong>{formatM2(totals.wrapArea)}</strong>
                </div>
                <div className="contra">
                  <span>Contra-Vision</span>
                  <strong>{formatM2(totals.contraArea)}</strong>
                </div>
                <div className="reflective">
                  <span>Reflective</span>
                  <strong>{formatM2(totals.reflectiveArea)}</strong>
                </div>
              </div>

              <div className="vinyl-summary-strip">
                <div>
                  <span>Labour hours</span>
                  <strong>{totals.labourHours.toFixed(1)}</strong>
                </div>
                <div>
                  <span>Total coverage</span>
                  <strong>{formatM2(totals.totalArea)}</strong>
                </div>
                <div>
                  <span>% coverage</span>
                  <strong>{formatPercent(totals.coverage)}</strong>
                </div>
              </div>

              {pricingEditable ? (
                <details className="vinyl-pricing-settings" onToggle={handlePricingSettingsToggle}>
                  <summary>Pricing settings</summary>
                  <p className="vinyl-pricing-context">
                    Editing pricing for <strong>{selectedTemplate.sizeName}</strong> ({selectedTemplate.exampleName}).
                  </p>
                  {pricingSyncStatus ? <p className="vinyl-pricing-context">{pricingSyncStatus}</p> : null}
                  <div className="smart-price-import">
                    <button
                      className="ghost-button smart-price-import-toggle"
                      type="button"
                      onClick={() => {
                        setSmartPriceImportOpen((isOpen) => !isOpen);
                        setSmartPriceImportStatus("");
                      }}
                    >
                      Import smart prices
                    </button>
                    {smartPriceImportOpen ? (
                      <div className="smart-price-import-panel">
                        <label>
                          <span>Paste pricing list</span>
                          <textarea
                            rows="7"
                            value={smartPriceImportText}
                            onChange={(event) => {
                              setSmartPriceImportText(event.target.value);
                              setSmartPriceImportStatus("");
                            }}
                            placeholder="Standard vinyl rate 84&#10;Wrap start rate 107&#10;Anchor 100% 3750"
                          />
                        </label>
                        <div className="smart-price-import-actions">
                          <button className="ghost-button" type="button" onClick={() => setSmartPriceImportText("")}>
                            Clear
                          </button>
                          <button className="primary-button" type="button" onClick={importSmartPrices}>
                            Import into {selectedTemplate.sizeName}
                          </button>
                        </div>
                        {smartPriceImportStatus ? <p className="smart-price-import-status">{smartPriceImportStatus}</p> : null}
                      </div>
                    ) : null}
                  </div>
                  <div className="vinyl-pricing-grid">
                    {renderPricingNumber("Standard vinyl rate", ["standardVinylRate"], 1, "Standard vinyl sell per m2. Raise it and all flat vinyl jobs rise.")}
                    {renderPricingNumber("Contra multiplier", ["materialMultipliers", "contra"], 0.05, "Multiplier for Contra-vision shapes. These stay standard vinyl, ignore wrap lines, and multiply the standard vinyl material sell.")}
                    {renderPricingNumber("Reflective multiplier", ["materialMultipliers", "reflective"], 0.05, "Multiplier for Reflective shapes. These stay standard vinyl, ignore wrap lines, and multiply the standard vinyl material sell.")}
                    {renderPricingNumber("Wrap start rate", ["wrapRateStart"], 1, "Starting wrap sell per m2 for very small wrap jobs. Raise it and small wrap jobs rise.")}
                    {renderPricingNumber("Wrap floor rate", ["wrapRateFloor"], 1, "Floor wrap sell per m2 for very large wrap jobs. Raise it and heavy partials and full wraps rise.")}
                    {renderPricingNumber("Wrap material taper", ["wrapRateTaper"], 1, "Controls how quickly wrap material rate falls as wrap coverage grows. Raise it and bigger wrap jobs get cheaper per m2 faster.")}
                    {renderPricingNumber("Small std hrs/m2", ["standardSmallHoursPerM2"], 0.05, "Controls how much labour small vinyl jobs carry before the cap. Raise it and small vinyl jobs rise.")}
                    {renderPricingNumber("Small std min hrs", ["standardSmallMinHours"], 0.05, "Small standard-vinyl labour minimum hours.")}
                    {renderPricingNumber("Small std max hrs", ["standardSmallMaxHours"], 0.05, "Small standard-vinyl labour maximum hours.")}
                    {renderPricingNumber("Large std hrs/m2", ["standardLargeHoursPerM2"], 0.05, "Labour hours per m2 for standard-vinyl jobs at 15 percent total coverage or more.")}
                    {renderPricingNumber("Wrap start hrs/m2", ["wrapLabourStartHoursPerM2"], 0.05, "Starting wrap labour hours per m2 for tiny wrap jobs.")}
                    {renderPricingNumber("Wrap floor hrs/m2", ["wrapLabourFloorHoursPerM2"], 0.05, "Floor wrap labour hours per m2 for very large wrap jobs.")}
                    {renderPricingNumber("Wrap labour taper", ["wrapLabourTaper"], 0.05, "Controls how fast wrap labour hours per m2 taper down as wrap coverage grows.")}
                    {renderPricingNumber("1 section factor", ["sectionFactors", "one"], 0.01, "Section factor for one connected wrap section. Raise it and single-section wrap jobs rise.")}
                    {renderPricingNumber("2-3 sections factor", ["sectionFactors", "twoToThree"], 0.01, "Section factor for 2 to 3 separate wrap sections. Raise it and broken-up wrap jobs rise.")}
                    {renderPricingNumber("4-5 sections factor", ["sectionFactors", "fourToFive"], 0.01, "Section factor for 4 to 5 separate wrap sections. Raise it and broken-up wrap jobs rise.")}
                    {renderPricingNumber("6+ sections factor", ["sectionFactors", "moreThanFive"], 0.01, "Section factor for 6 or more separate wrap sections. Raise it and broken-up wrap jobs rise.")}
                    {renderPricingNumber("Normal wrap difficulty", ["difficultyFactors", "normal_wrap_curve"], 0.01, "Difficulty factor for normal wrap curves. Raise it and curved wrap zones rise.")}
                    {renderPricingNumber("Labour sell/hr", ["labourSellRate"], 1, "Labour sell per hour. Raise it and labour-heavy jobs rise.")}
                    {renderPricingNumber("Anchor 0%", ["marketAnchors", "c0"], 1, "Market anchor at 0 percent total coverage.")}
                    {renderPricingNumber("Anchor 5%", ["marketAnchors", "c05"], 1, "Market anchor at 5 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 10%", ["marketAnchors", "c10"], 1, "Market anchor at 10 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 15%", ["marketAnchors", "c15"], 1, "Market anchor at 15 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 22%", ["marketAnchors", "c22"], 1, "Market anchor at 22 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 35%", ["marketAnchors", "c35"], 1, "Market anchor at 35 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 55%", ["marketAnchors", "c55"], 1, "Market anchor at 55 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 85%", ["marketAnchors", "c85"], 1, "Market anchor at 85 percent total coverage. Raise it and jobs around this coverage rise.")}
                    {renderPricingNumber("Anchor 100%", ["marketAnchors", "c100"], 1, "Market anchor at 100 percent total coverage. Raise it and full-coverage jobs rise.")}
                    {renderPricingNumber("No-wrap calc weight", ["blendWeights", "noWrap", "calculated"], 0.05, "Calculated-price weight for jobs with no wrap film. More calculated weight makes technical cost drive the final price more.")}
                    {renderPricingNumber("No-wrap anchor weight", ["blendWeights", "noWrap", "anchor"], 0.05, "Anchor weight for jobs with no wrap film. More anchor weight keeps prices inside a tighter market band.")}
                    {renderPricingNumber("Wrap <35 calc weight", ["blendWeights", "wrapUnder35", "calculated"], 0.05, "Calculated-price weight for wrap jobs below 35 percent total coverage.")}
                    {renderPricingNumber("Wrap <35 anchor weight", ["blendWeights", "wrapUnder35", "anchor"], 0.05, "Anchor weight for wrap jobs below 35 percent total coverage.")}
                    {renderPricingNumber("Wrap <70 calc weight", ["blendWeights", "wrapUnder70", "calculated"], 0.05, "Calculated-price weight for wrap jobs from 35 percent to under 70 percent total coverage.")}
                    {renderPricingNumber("Wrap <70 anchor weight", ["blendWeights", "wrapUnder70", "anchor"], 0.05, "Anchor weight for wrap jobs from 35 percent to under 70 percent total coverage.")}
                    {renderPricingNumber("Wrap 70%+ calc weight", ["blendWeights", "wrapFull", "calculated"], 0.05, "Calculated-price weight for wrap jobs at 70 percent total coverage or more.")}
                    {renderPricingNumber("Wrap 70%+ anchor weight", ["blendWeights", "wrapFull", "anchor"], 0.05, "Anchor weight for wrap jobs at 70 percent total coverage or more.")}
                    {renderPricingNumber("Absolute minimum", ["minPrice"], 1, "Absolute minimum for any job. Stops very small jobs dropping too low.")}
                    {renderPricingNumber("Any wrap minimum", ["minAnyWrapPrice"], 1, "Minimum when any wrap film is present.")}
                    {renderPricingNumber("Partial wrap minimum", ["minPartialWrapPrice"], 1, "Minimum when total coverage is 15 percent or more and wrap is present.")}
                    {renderPricingNumber("Full wrap minimum", ["minFullWrapPrice"], 1, "Minimum when total coverage is 85 percent or more.")}
                  </div>
                  <div className="vinyl-pricing-actions">
                    <button className="ghost-button" type="button" onClick={resetPricingDraft}>
                      Revert
                    </button>
                    <button className="ghost-button" type="button" onClick={restoreDefaultPricingSettings}>
                      Defaults
                    </button>
                    <button className="primary-button" type="button" onClick={savePricingSettings}>
                      Save
                    </button>
                  </div>
                  <div className="price-training-bank">
                    <div>
                      <p className="eyebrow">Smart calibration</p>
                      <h4>Training bank</h4>
                      <p>
                        Draw the boxes, enter the price you would actually charge, then add it to the bank. The suggestion
                        uses your saved examples for this vehicle size only.
                      </p>
                    </div>
                    <div className="price-training-add">
                      <label>
                        <span>Your price ex VAT</span>
                        <input
                          type="number"
                          min="0"
                          step="50"
                          value={trainingTargetPrice}
                          onChange={(event) => {
                            setTrainingTargetPrice(event.target.value);
                            setTrainingStatus("");
                          }}
                          placeholder="e.g. 1400"
                        />
                      </label>
                      <button className="ghost-button" type="button" onClick={addTrainingExample}>
                        Add current drawing to bank
                      </button>
                    </div>
                    {selectedTemplateTrainingBank.length ? (
                      <div className="price-training-list">
                        {selectedTemplateTrainingBank.slice(0, 8).map((sample) => {
                          const sampleEstimate = calculateEstimateFromMetrics(sample.metrics, pricingDraftSettings);
                          return (
                            <div className="price-training-row" key={sample.id}>
                              <div>
                                <strong>{currencyFormatter.format(sample.targetPrice)}</strong>
                                <span>
                                  {formatM2(sample.metrics.totalArea)} total, {formatM2(sample.metrics.wrapArea)} wrap
                                </span>
                              </div>
                              <span>{currencyFormatter.format(sampleEstimate)} now</span>
                              <button className="text-button danger" type="button" onClick={() => removeTrainingExample(sample.id)}>
                                Delete
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <p className="price-training-empty">No banked examples for this vehicle yet.</p>
                    )}
                    <div className="price-training-actions">
                      <button className="ghost-button" type="button" onClick={calculateTrainingSuggestion}>
                        Work out suggested variables
                      </button>
                      <button className="primary-button" type="button" onClick={applyTrainingSuggestion} disabled={!trainingSuggestion}>
                        Apply suggestion
                      </button>
                    </div>
                    {trainingSuggestion ? (
                      <div className="price-training-suggestion">
                        <strong>{trainingSuggestion.label}</strong>
                        <span>
                          Average error: {Math.round(trainingSuggestion.beforeScore * 100)}% now to{" "}
                          {Math.round(trainingSuggestion.score * 100)}% suggested
                        </span>
                      </div>
                    ) : null}
                    {trainingStatus ? <p className="smart-price-import-status">{trainingStatus}</p> : null}
                  </div>
                </details>
              ) : null}

            </aside>
          </div>
        </section>

        {pdfExportOpen ? (
          <div className="modal-backdrop" onMouseDown={(event) => {
            if (event.target === event.currentTarget) setPdfExportOpen(false);
          }}>
            <form className="vehicle-pdf-modal" onSubmit={exportVehiclePdf}>
              <div className="vehicle-pdf-modal-header">
                <div>
                  <p className="eyebrow">Vehicle PDF</p>
                  <h2>Export proposal</h2>
                </div>
                <button className="icon-button" type="button" onClick={() => setPdfExportOpen(false)}>
                  x
                </button>
              </div>
              {pdfExportError ? <div className="flash error">{pdfExportError}</div> : null}
              <label>
                Customer name
                <input
                  type="text"
                  value={pdfExportForm.customerName}
                  onChange={(event) => setPdfExportForm((current) => ({ ...current, customerName: event.target.value }))}
                  required
                  autoFocus
                />
              </label>
              <label>
                Notes
                <textarea
                  rows="6"
                  value={pdfExportForm.notes}
                  onChange={(event) => setPdfExportForm((current) => ({ ...current, notes: event.target.value }))}
                  placeholder="Anything you want shown at the bottom of the detail box..."
                />
              </label>
              <div className="modal-actions">
                <button className="ghost-button" type="button" onClick={() => setPdfExportOpen(false)} disabled={pdfExporting}>
                  Cancel
                </button>
                <button className="primary-button" type="submit" disabled={pdfExporting}>
                  {pdfExporting ? "Creating..." : "Create PDF"}
                </button>
              </div>
            </form>
          </div>
        ) : null}
      </div>
    </div>
  );
}
export {VinylEstimatorPage as default,getDefaultVehiclePricingSettingsByTemplate,VEHICLE_TEMPLATE_OPTIONS};
