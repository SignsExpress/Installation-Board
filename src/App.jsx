import { useEffect, useMemo, useRef, useState } from "react";
import InstallerDirectoryHost from "./installer/InstallerDirectoryHostV2";

const JOB_TYPES = [
  { value: "Install", colorClass: "job-type-install" },
  { value: "Vehicle", colorClass: "job-type-vehicle" },
  { value: "Delivery", colorClass: "job-type-delivery" },
  { value: "Subcontractor", colorClass: "job-type-subcontractor" },
  { value: "Signs Express", colorClass: "job-type-signs-express" },
  { value: "Survey", colorClass: "job-type-survey" },
  { value: "Other", colorClass: "job-type-other" }
];

const INSTALLER_OPTIONS = [
  { value: "MC", colorClass: "installer-mc" },
  { value: "KC", colorClass: "installer-kc" },
  { value: "ED", colorClass: "installer-ed" },
  { value: "KW", colorClass: "installer-kw" },
  { value: "PM", colorClass: "installer-pm" },
  { value: "MR", colorClass: "installer-mr" },
  { value: "Custom", colorClass: "installer-custom" }
];

const PERMISSION_OPTIONS = [
  { value: "admin", label: "Admin" },
  { value: "user", label: "User" },
  { value: "none", label: "No Access" }
];

const HOLIDAY_STAFF = [
  { code: "MR", person: "Matt R", fullName: "Matt Rutlidge", colorClass: "holiday-person-black", birthDate: "" },
  { code: "DD", person: "Dawn D", fullName: "Dawn Dewhurst", colorClass: "holiday-person-black", birthDate: "" },
  { code: "TVB", person: "Tom V-B", fullName: "Tom Van-Boyd", colorClass: "holiday-person-black", birthDate: "" },
  { code: "AH", person: "Amber H", fullName: "Amber Hardman", colorClass: "holiday-person-black", birthDate: "" },
  { code: "ED", person: "Eddy D'A", fullName: "Eddy D'Antonio", colorClass: "holiday-person-black", birthDate: "" },
  { code: "PM", person: "Paul M", fullName: "Paul Morris", colorClass: "holiday-person-green", birthDate: "" },
  { code: "KW", person: "Kyle W", fullName: "Kyle Wright", colorClass: "holiday-person-green", birthDate: "" },
  { code: "MC", person: "Matt C", fullName: "Matt Carroll", colorClass: "holiday-person-red", birthDate: "" },
  { code: "KC", person: "Keilan C", fullName: "Keilan Curtis", colorClass: "holiday-person-red", birthDate: "" }
];
const HOLIDAY_PERSON_COLORS = Object.fromEntries(HOLIDAY_STAFF.map((entry) => [entry.person, entry.colorClass]));
const UNSCHEDULED_DROP_ZONE = "__unscheduled__";

const EMPTY_FORM = {
  id: "",
  date: "",
  orderReference: "",
  customerName: "",
  description: "",
  contact: "",
  number: "",
  address: "",
  installers: [],
  customInstaller: "",
  jobType: "Install",
  customJobType: "",
  isPlaceholder: false,
  notes: ""
};

const EMPTY_HOLIDAY_REQUEST_FORM = {
  person: "",
  startDate: "",
  endDate: "",
  duration: "Full Day",
  notes: ""
};

const EMPTY_HOLIDAY_CANCEL_FORM = {
  requestId: "",
  notes: ""
};

const EMPTY_HOLIDAY_EVENT_FORM = {
  id: "",
  date: "",
  title: ""
};

const EMPTY_ATTENDANCE_NOTE_FORM = {
  date: "",
  note: ""
};

const RAMS_ACTIVITY_OPTIONS = [
  { value: "internal", label: "Internal signs / wall graphics" },
  { value: "external", label: "External signs / fascia" },
  { value: "vehicle", label: "Vehicle graphics" },
  { value: "window", label: "Window graphics" },
  { value: "survey", label: "Survey / measurements only" }
];

const RAMS_ACCESS_OPTIONS = [
  { value: "ground", label: "Ground level" },
  { value: "steps", label: "Steps / podium" },
  { value: "ladders", label: "Ladders" },
  { value: "mewp", label: "MEWP / powered access" }
];

const RAMS_WORK_AREA_OPTIONS = [
  { value: "quiet", label: "Quiet controlled area" },
  { value: "public", label: "Public / occupied area" },
  { value: "traffic", label: "Near vehicles / deliveries" },
  { value: "construction", label: "Construction site" }
];

const RAMS_TOOL_OPTIONS = [
  { value: "hand-tools", label: "Hand tools" },
  { value: "power-tools", label: "Power tools / drilling" },
  { value: "adhesives", label: "Adhesives / cleaners" },
  { value: "lifting", label: "Heavy or awkward lifting" },
  { value: "electrical", label: "Electrical isolation nearby" }
];

const RAMS_DEFAULT_QUESTIONS = {
  jobId: "",
  activity: "external",
  access: "ground",
  workArea: "quiet",
  tools: ["hand-tools"],
  operatives: "2",
  duration: "1 day",
  welfare: "Client welfare facilities to be confirmed at induction.",
  emergency: "Follow site emergency arrangements and report incidents to the site contact and Signs Express.",
  notes: ""
};

const RAMS_STANDARD_RISK_CARDS = {
  accessEgress: {
    title: "Access & Egress",
    type: "Risk",
    trigger: "Always included",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 3,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 1,
    responsibility: "Matt Carroll",
    controlMeasure: "An agreed route to and from place of work will be kept to. Traffic movement will be done in a safe and orderly manner in conjunction with site manager's instruction.",
    content: ["An agreed route to and from place of work will be kept to. Traffic movement will be done in a safe and orderly manner in conjunction with site manager's instruction."]
  },
  injuryIncident: {
    title: "Risk of Injury or Incident",
    type: "Risk",
    trigger: "Always included",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 3,
    initialConsequence: 3,
    residualLikelihood: 1,
    residualConsequence: 2,
    responsibility: "Matt Carroll",
    controlMeasure: "All fitters undergo training for work positioning and correct use of tools and equipment. In compliance with current guidelines, a safety rescue plan is discussed and decided prior to any work being carried out. First aid facilities available. An accident report will be filed after any incident.",
    content: ["All fitters undergo training for work positioning and correct use of tools and equipment.", "Safety rescue plan to be discussed before works start.", "First aid facilities available and accident report completed after any incident."]
  },
  musculoskeletal: {
    title: "Musculoskeletal Disorders",
    type: "Risk",
    trigger: "Heavy or awkward lifting selected",
    whoAtRisk: "Employees",
    initialLikelihood: 2,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 2,
    responsibility: "Matt Carroll",
    controlMeasure: "All fitters undergo training for correct manual handling techniques which are to be adhered to at all times.",
    content: ["All fitters undergo training for correct manual handling techniques which are to be adhered to at all times."]
  },
  ppeUnsuitable: {
    title: "Injury to Fitters Due to PPE Unsuitable for the Environment Being Worked In",
    type: "Risk",
    trigger: "Always included",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 3,
    initialConsequence: 3,
    residualLikelihood: 1,
    residualConsequence: 2,
    responsibility: "Matt Carroll",
    controlMeasure: "PPE is provided and worn by fitters. PPE will be appropriate not only to the job being undertaken but also meeting the requirements of the site as a whole.",
    content: ["PPE is provided and worn by fitters.", "PPE will be appropriate to the job and the requirements of the site as a whole."]
  },
  poorLighting: {
    title: "Risk of Injury Due to Poor Lighting",
    type: "Risk",
    trigger: "Internal or low light works",
    whoAtRisk: "Employees",
    initialLikelihood: 2,
    initialConsequence: 3,
    residualLikelihood: 1,
    residualConsequence: 2,
    responsibility: "Matt Carroll",
    controlMeasure: "External work to be carried out in daylight. Site lighting to be used for internal work where required.",
    content: ["External work to be carried out in daylight.", "Site lighting to be used for internal work where required."]
  },
  height: {
    title: "Falls from Height",
    type: "Risk",
    trigger: "Steps, ladders or MEWP",
    whoAtRisk: "Employees",
    initialLikelihood: 2,
    initialConsequence: 4,
    residualLikelihood: 1,
    residualConsequence: 4,
    responsibility: "Matt Carroll",
    controlMeasure: "All work to be carried out by competent fitters who have undertaken ladder and step-up training. Fitters will not erect scaffolds and will supervise any other colleagues when erecting or using scaffold.",
    content: ["Use the lowest-risk access method suitable for the task and inspect access equipment before use.", "Competent fitters only to work at height and avoid overreaching.", "Do not work at height in unsafe weather, poor ground conditions or uncontrolled public areas."]
  },
  public: {
    title: "Injury to the General Public During Course of Work",
    type: "Risk",
    trigger: "Public or occupied area",
    whoAtRisk: "Third parties",
    initialLikelihood: 2,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 1,
    responsibility: "Matt Carroll",
    controlMeasure: "Report all hazardous activity to site management. Ensure a fitter is present with the ability to communicate with workers and emergency services.",
    content: ["Set a clear exclusion zone using cones, barriers or signage suitable for the work area.", "Keep tools, materials and waste within the controlled area and maintain safe pedestrian routes.", "Suspend work if the exclusion zone cannot be maintained."]
  },
  fallingObjects: {
    title: "Falling Objects",
    type: "Risk",
    trigger: "Steps, ladders or MEWP",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 2,
    initialConsequence: 5,
    residualLikelihood: 1,
    residualConsequence: 5,
    responsibility: "Matt Carroll",
    controlMeasure: "Area cordoned off with cones and safety barriers. Provision to be made to allow building occupants to escape if blocking escape route. All small tools and materials will be carried in a bag or on a belt. All large tools should have carry straps. All scaffold to have toe boards.",
    content: ["Area cordoned off with cones and safety barriers.", "Small tools and materials to be carried in a bag or on a belt.", "Large tools should have carry straps and scaffold should have toe boards."]
  },
  electricalEquipment: {
    title: "Electric Shock from Equipment",
    type: "Risk",
    trigger: "Power tools / drilling selected",
    whoAtRisk: "Employees",
    initialLikelihood: 0,
    initialConsequence: 0,
    residualLikelihood: 0,
    residualConsequence: 0,
    responsibility: "Matt Carroll",
    controlMeasure: "No electrical equipment will be used and no work will be done on fixed electrical system. All drills etc are battery operated or powered by 110v generator.",
    content: ["No electrical equipment will be used and no work will be done on fixed electrical system.", "All drills etc are battery operated or powered by 110v generator."]
  },
  equipmentFailure: {
    title: "Equipment Failure",
    type: "Risk",
    trigger: "Power tools / drilling selected",
    whoAtRisk: "Employees",
    initialLikelihood: 2,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 1,
    responsibility: "Matt Carroll",
    controlMeasure: "Ensure equipment is routinely checked. Routine inspection of all equipment prior to commencement of job. Routine testing and maintenance schedule established and in operation. All electrical equipment has been PAT tested.",
    content: ["Ensure equipment is routinely checked.", "Routine inspection of all equipment prior to commencement of job.", "Routine testing and maintenance schedule established and in operation."]
  },
  fittersFatigue: {
    title: "Fitters Feeling Fatigue",
    type: "Risk",
    trigger: "Always included",
    whoAtRisk: "Employees",
    initialLikelihood: 1,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 1,
    responsibility: "Matt Carroll",
    controlMeasure: "The supervisor will keep in contact with the fitters regularly. If they become tired another fitter will take over.",
    content: ["The supervisor will keep in contact with the fitters regularly.", "If they become tired another fitter will take over."]
  },
  slipsTrips: {
    title: "Slips, Trips & Falls",
    type: "Risk",
    trigger: "Always included",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 2,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 2,
    responsibility: "Matt Carroll",
    controlMeasure: "Area to be inspected for hazards and any objects removed if possible. If not, advise others. All areas of work to be coned off.",
    content: ["Area to be inspected for hazards and any objects removed if possible.", "If hazards cannot be removed, advise others.", "All areas of work to be coned off."]
  },
  badWeather: {
    title: "Bad Weather",
    type: "Risk",
    trigger: "External works",
    whoAtRisk: "Employees",
    initialLikelihood: 1,
    initialConsequence: 1,
    residualLikelihood: 1,
    residualConsequence: 1,
    responsibility: "Matt Carroll",
    controlMeasure: "Work to be called off if weather conditions become unsafe.",
    content: ["Work to be called off if weather conditions become unsafe."]
  },
  diggingHoles: {
    title: "Digging Holes",
    type: "Risk",
    trigger: "Ground works",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 1,
    initialConsequence: 2,
    residualLikelihood: 1,
    residualConsequence: 1,
    responsibility: "Matt Carroll",
    controlMeasure: "Area to be CAT scanned prior to any digging.",
    content: ["Area to be CAT scanned prior to any digging."]
  },
  windowBreakage: {
    title: "Accidental Breakage of Windows",
    type: "Risk",
    trigger: "Window graphics selected",
    whoAtRisk: "Employees\nThird parties",
    initialLikelihood: 2,
    initialConsequence: 3,
    residualLikelihood: 1,
    residualConsequence: 3,
    responsibility: "Matt Carroll",
    controlMeasure: "Care to be taken to ensure that materials and tools are kept away from any glass.",
    content: ["Care to be taken to ensure that materials and tools are kept away from any glass."]
  },
  emergencyIncidents: {
    title: "Emergency Incidents on Site",
    type: "Risk",
    trigger: "Always included",
    whoAtRisk: "Employees",
    initialLikelihood: 2,
    initialConsequence: 3,
    residualLikelihood: 1,
    residualConsequence: 3,
    responsibility: "Matt Carroll",
    controlMeasure: "Procedures for emergency incidents on site to be agreed with Signs Express personnel and site manager prior to operation.",
    content: ["Procedures for emergency incidents on site to be agreed with Signs Express personnel and site manager prior to operation."]
  }
};

const RAMS_CARD_LIBRARY = {
  induction: {
    title: "Arrive, Sign In and Confirm Site Controls",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "All operatives sign in, attend site induction where required and follow client or principal contractor rules.",
      "Confirm working area, emergency arrangements, welfare, permit requirements and any live restrictions before work starts.",
      "Stop work and report to the site contact if conditions differ from the agreed RAMS."
    ]
  },
  siteSetup: {
    title: "Set Up Working Area",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Low",
    residualRisk: "Low",
    content: [
      "Park and unload in the agreed area without blocking fire routes, access routes or live work areas.",
      "Move materials to the work face and keep packaging, tools and fixings organised.",
      "Create a neat working zone before installation starts so the task can be completed without unnecessary movement around site."
    ]
  },
  surveyCheck: {
    title: "Check Artwork, Position and Substrate",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Low",
    residualRisk: "Low",
    content: [
      "Confirm location, sizes, artwork orientation and fixing positions against the job details before installation.",
      "Check the substrate is suitable, clean, dry and ready to accept fixings, vinyl or adhesive.",
      "Raise any mismatch, damaged surface or access restriction before committing materials."
    ]
  },
  accessSetup: {
    title: "Prepare Access Equipment",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Choose the agreed access method for the task and check it is suitable for the surface, height and duration.",
      "Inspect steps, podiums, ladders or platform equipment before use.",
      "Keep access equipment within the controlled work area and reposition it instead of overreaching."
    ]
  },
  installSequence: {
    title: "Install Signage or Graphics",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Offer up, align and temporarily position materials before final fixing or application.",
      "Use suitable fixings, tapes, adhesives or application methods for the surface and product type.",
      "Keep blades, drills and small tools controlled while working and avoid leaving loose items at height."
    ]
  },
  qualityCheck: {
    title: "Quality Check and Make Good",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Low",
    residualRisk: "Low",
    content: [
      "Check alignment, finish, adhesion, fixings and visible marks before leaving the work face.",
      "Clean down the installed area where required and remove application marks or debris.",
      "Photograph completed works where required for job records or client sign-off."
    ]
  },
  public: {
    title: "Public, Staff and Third-Party Interface",
    type: "Risk",
    trigger: "Public or occupied area",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Set a clear exclusion zone using cones, barriers or signage suitable for the work area.",
      "Keep tools, materials and waste within the controlled area and maintain safe pedestrian routes.",
      "Suspend work if the exclusion zone cannot be maintained."
    ]
  },
  slipsTrips: {
    title: "Slips, Trips and Housekeeping",
    type: "Risk",
    trigger: "Always included",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Keep walkways, access routes and the working area clear of tools, packaging, trailing leads and loose materials.",
      "Clean up debris as work progresses and remove waste from site or place it in the agreed waste area.",
      "Do not leave materials where they could fall, blow away or create a trip hazard."
    ]
  },
  traffic: {
    title: "Vehicle Movement and Deliveries",
    type: "Risk",
    trigger: "Vehicle movement nearby",
    initialRisk: "High",
    residualRisk: "Low",
    content: [
      "Agree safe parking, unloading and working positions with site before work starts.",
      "Use a banksman where visibility is restricted or materials are moved near vehicle routes.",
      "Wear hi-vis PPE and do not work in live traffic routes without suitable segregation."
    ]
  },
  construction: {
    title: "Construction Site Controls",
    type: "Method",
    trigger: "Construction site",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Follow principal contractor rules, daily briefings and permit systems.",
      "Wear site-required PPE as a minimum and keep the work area tidy.",
      "Coordinate with other trades before drilling, lifting or closing access routes."
    ]
  },
  height: {
    title: "Working at Height",
    type: "Risk",
    trigger: "Steps, ladders or MEWP",
    initialRisk: "High",
    residualRisk: "Low",
    content: [
      "Use the lowest-risk access method suitable for the task and inspect access equipment before use.",
      "Maintain three points of contact where applicable and avoid overreaching.",
      "Do not work at height in unsafe weather, poor ground conditions or uncontrolled public areas."
    ]
  },
  ladders: {
    title: "Ladders and Step Ladders",
    type: "Method",
    trigger: "Ladders selected",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Use ladders only for short-duration light work where a safer platform is not reasonably practicable.",
      "Set ladders on firm, level ground and secure or foot them where required.",
      "Keep both hands free while climbing and raise materials separately where needed."
    ]
  },
  mewp: {
    title: "MEWP / Powered Access",
    type: "Method",
    trigger: "MEWP selected",
    initialRisk: "High",
    residualRisk: "Low",
    content: [
      "MEWP to be operated only by trained, authorised operatives.",
      "Complete pre-use checks, confirm ground conditions and establish an exclusion zone.",
      "Wear harness and lanyard where required by the platform type and site rules."
    ]
  },
  tools: {
    title: "Hand Tools, Cutting and Drilling",
    type: "Risk",
    trigger: "Tools selected",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Use the correct tool for the task and check tools before use.",
      "Wear eye protection when drilling, cutting or fixing.",
      "Control dust and debris, and check for hidden services before drilling."
    ]
  },
  substances: {
    title: "Adhesives, Cleaners and Substances",
    type: "COSHH",
    trigger: "Adhesives / cleaners selected",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Use only required quantities and follow product safety data sheets.",
      "Maintain ventilation, avoid ignition sources and wear suitable gloves where needed.",
      "Store containers securely and remove waste from site."
    ]
  },
  lifting: {
    title: "Manual Handling and Material Movement",
    type: "Risk",
    trigger: "Heavy or awkward lifting selected",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Assess size, weight, route and fixing position before lifting.",
      "Use team lifts, trolleys or mechanical aids for heavy or awkward items.",
      "Keep routes clear and avoid twisting while carrying materials."
    ]
  },
  electrical: {
    title: "Electrical Services and Isolation",
    type: "Risk",
    trigger: "Electrical work nearby",
    initialRisk: "High",
    residualRisk: "Low",
    content: [
      "Signs Express operatives must not work on live electrical systems unless specifically competent and authorised.",
      "Confirm isolation or safe routing before drilling near lighting, cables or illuminated sign supplies.",
      "Stop work if unknown services are found."
    ]
  },
  electricalEquipment: {
    title: "Electric Shock from Equipment",
    type: "Risk",
    trigger: "Power tools / drilling selected",
    initialRisk: "High",
    residualRisk: "Low",
    content: [
      "Use battery-powered tools where reasonably practicable and check tools, chargers and leads before use.",
      "Do not use damaged electrical equipment, exposed cables or improvised connections.",
      "Keep electrical equipment away from wet surfaces and route leads to avoid damage or trip hazards."
    ]
  },
  vehicle: {
    title: "Vehicle Graphics Installation",
    type: "Method",
    trigger: "Vehicle graphics selected",
    initialRisk: "Low",
    residualRisk: "Low",
    content: [
      "Confirm vehicle is parked safely with adequate space around the work area.",
      "Clean and prepare surfaces using approved products before applying graphics.",
      "Keep blades controlled and dispose of backing paper, application tape and waste safely."
    ]
  },
  weather: {
    title: "Weather and External Conditions",
    type: "Risk",
    trigger: "External works",
    initialRisk: "Medium",
    residualRisk: "Low",
    content: [
      "Check weather conditions before external installation starts.",
      "Do not install in high wind, heavy rain, icy conditions or temperatures unsuitable for the materials.",
      "Secure loose materials and postpone work if conditions become unsafe."
    ]
  },
  completion: {
    title: "Completion, Handover and Waste",
    type: "Method",
    trigger: "Always included",
    initialRisk: "Low",
    residualRisk: "Low",
    content: [
      "Inspect finished work, remove waste and leave the area tidy.",
      "Return access routes and barriers to the agreed arrangement.",
      "Report completion, issues or snagging items to the site contact."
    ]
  }
};

const RAMS_BASE_CARD_IDS = [
  "induction",
  "accessEgress",
  "injuryIncident",
  "slipsTrips",
  "ppeUnsuitable",
  "fittersFatigue",
  "emergencyIncidents",
  "tools",
  "lifting",
  "siteSetup",
  "surveyCheck",
  "accessSetup",
  "installSequence",
  "qualityCheck",
  "completion"
];

const RAMS_LOGIC_STORAGE_KEY = "rams-builder-logic-v1";

const RAMS_DEFAULT_LOGIC = {
  optionGroups: [
    {
      key: "activity",
      label: "Work type",
      input: "buttons",
      multi: false,
      options: RAMS_ACTIVITY_OPTIONS.map((option) => ({
        ...option,
        cardIds:
          option.value === "external"
            ? ["weather"]
            : option.value === "window"
              ? ["weather"]
              : option.value === "vehicle"
                ? ["vehicle"]
                : []
      }))
    },
    {
      key: "access",
      label: "Access method",
      input: "select",
      multi: false,
      options: RAMS_ACCESS_OPTIONS.map((option) => ({
        ...option,
        cardIds:
          option.value === "steps"
            ? ["height", "fallingObjects"]
            : option.value === "ladders"
              ? ["height", "fallingObjects", "ladders"]
              : option.value === "mewp"
                ? ["height", "fallingObjects", "mewp"]
                : []
      }))
    },
    {
      key: "workArea",
      label: "Work area",
      input: "select",
      multi: false,
      options: RAMS_WORK_AREA_OPTIONS.map((option) => ({
        ...option,
        cardIds:
          option.value === "public"
            ? ["public"]
            : option.value === "traffic"
              ? ["traffic"]
              : option.value === "construction"
                ? ["construction"]
                : []
      }))
    },
    {
      key: "tools",
      label: "Tools and conditions",
      input: "checkboxes",
      multi: true,
      options: RAMS_TOOL_OPTIONS.map((option) => ({
        ...option,
        cardIds:
          option.value === "power-tools"
            ? ["tools", "electricalEquipment", "equipmentFailure"]
            : option.value === "adhesives"
              ? ["substances"]
            : option.value === "lifting"
                ? ["lifting", "musculoskeletal"]
                : option.value === "electrical"
                  ? ["electrical"]
                  : []
      }))
    }
  ],
  cards: {
    ...RAMS_CARD_LIBRARY,
    ...RAMS_STANDARD_RISK_CARDS
  },
  baseCardIds: RAMS_BASE_CARD_IDS
};

const ATTENDANCE_WEEKDAYS = [
  ["monday", "Mon"],
  ["tuesday", "Tue"],
  ["wednesday", "Wed"],
  ["thursday", "Thu"],
  ["friday", "Fri"]
];

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

const VEHICLE_PRICING_STORAGE_KEY = "vehicle-pricing-settings-formula-v2";
const VEHICLE_PRICE_TRAINING_BANK_STORAGE_KEY = "vehicle-price-training-bank-v1";

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

function getStoredVehiclePricingSettingsByTemplate() {
  const defaultSettings = getDefaultVehiclePricingSettingsByTemplate();
  if (typeof window === "undefined") return defaultSettings;
  try {
    const storedSettings = window.localStorage.getItem(VEHICLE_PRICING_STORAGE_KEY);
    if (!storedSettings) return defaultSettings;
    const parsedSettings = JSON.parse(storedSettings);

    if (looksLikeVehiclePricingSettings(parsedSettings)) return defaultSettings;

    return Object.fromEntries(
      VEHICLE_TEMPLATE_OPTIONS.map((template) => {
        const storedTemplateSettings = parsedSettings?.[template.id] || {};
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

function getLocalTodayIso() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(new Date());
}

function parseIsoDate(isoDate) {
  const [year, month, day] = String(isoDate || "").split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(Date.UTC(year, month - 1, day));
}

function toIsoDate(date) {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "UTC",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(date);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setUTCDate(next.getUTCDate() + days);
  return next;
}

function addMonths(date, months) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
}

function getStartOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function getEndOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
}

function getHolidayStaffEntry(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return (
    HOLIDAY_STAFF.find((entry) => entry.fullName.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.person.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.code.toLowerCase() === normalized) ||
    null
  );
}

function getHolidayStaffPersonForUser(user) {
  return getHolidayStaffEntry(user?.displayName)?.person || "";
}

function getHolidayStaffIdentityKey(value) {
  const match = getHolidayStaffEntry(value);
  if (match?.code) return match.code.toLowerCase();
  return String(value || "").trim().toLowerCase();
}

function getHolidayDisplayToken(person) {
  const entry = getHolidayStaffEntry(person);
  if (entry) return entry.code;
  const value = String(person || "").trim();
  if (value && !value.includes(" ") && value.length <= 4) return value.toUpperCase();
  return toInitials(value);
}

function normalizeHolidayStaffEntries(entries) {
  if (!Array.isArray(entries) || !entries.length) return HOLIDAY_STAFF;
  return entries.map((entry) => ({
    ...entry,
    fullName: entry.fullName || entry.name || entry.person || ""
  }));
}

function toAllowanceNumber(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
}

function normalizeAttendanceDraft(profile) {
  const base = {
    mode: "required",
    contractedHours: Object.fromEntries(
      ATTENDANCE_WEEKDAYS.map(([dayKey]) => [dayKey, { in: "", out: "", off: false }])
    )
  };
  const nextMode = ["required", "fixed", "exempt"].includes(String(profile?.mode || "").trim().toLowerCase())
    ? String(profile.mode).trim().toLowerCase()
    : "required";

  for (const [dayKey] of ATTENDANCE_WEEKDAYS) {
    base.contractedHours[dayKey] = {
      in: String(profile?.contractedHours?.[dayKey]?.in || "").trim(),
      out: String(profile?.contractedHours?.[dayKey]?.out || "").trim(),
      off: Boolean(profile?.contractedHours?.[dayKey]?.off)
    };
  }

  return {
    mode: nextMode,
    contractedHours: base.contractedHours
  };
}

function getHolidayAllowanceSummary(entry) {
  const standardEntitlement = toAllowanceNumber(entry.standardEntitlement);
  const extraServiceDays = toAllowanceNumber(entry.extraServiceDays);
  const christmasDays = toAllowanceNumber(entry.christmasDays);
  const bankHolidayDays = toAllowanceNumber(entry.bankHolidayDays);
  const bookedDays = toAllowanceNumber(entry.bookedDays);
  const unpaidDaysBooked = toAllowanceNumber(entry.unpaidDaysBooked);
  const workDaysPerWeek = toAllowanceNumber(entry.workDaysPerWeek);
  const prorataAllowance = standardEntitlement + extraServiceDays;
  const daysLeft = prorataAllowance - christmasDays - bankHolidayDays - bookedDays;

  return {
    ...entry,
    workDaysPerWeek,
    standardEntitlement,
    extraServiceDays,
    christmasDays,
    bankHolidayDays,
    bookedDays,
    unpaidDaysBooked,
    prorataAllowance,
    daysLeft
  };
}

function formatHolidayBirthday(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return "-";
  return parsed.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: "UTC"
  });
}

function isBirthdayHoliday(entry) {
  return String(entry?.type || "").trim().toLowerCase() === "birthday";
}

function getCurrentHolidayYearStart(anchorIsoDate = getLocalTodayIso()) {
  const anchor = parseIsoDate(anchorIsoDate) || parseIsoDate(getLocalTodayIso());
  if (!anchor) return new Date().getUTCFullYear();
  return anchor.getUTCMonth() >= 1 ? anchor.getUTCFullYear() : anchor.getUTCFullYear() - 1;
}

function getHolidayYearLabel(yearStart) {
  return `${yearStart}-${String(yearStart + 1).slice(-2)}`;
}

function nthWeekdayOfMonth(year, month, weekday, occurrence) {
  const first = new Date(Date.UTC(year, month, 1));
  const firstWeekday = first.getUTCDay();
  const delta = (weekday - firstWeekday + 7) % 7;
  return new Date(Date.UTC(year, month, 1 + delta + (occurrence - 1) * 7));
}

function lastWeekdayOfMonth(year, month, weekday) {
  const last = new Date(Date.UTC(year, month + 1, 0));
  const lastWeekday = last.getUTCDay();
  const delta = (lastWeekday - weekday + 7) % 7;
  return new Date(Date.UTC(year, month + 1, last.getUTCDate() - delta));
}

function getEasterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(year, month - 1, day));
}

function applySubstituteHoliday(date) {
  const weekday = date.getUTCDay();
  if (weekday === 6) return addDays(date, 2);
  if (weekday === 0) return addDays(date, 1);
  return date;
}

function getUkBankHolidays(year) {
  const easterSunday = getEasterSunday(year);
  const goodFriday = addDays(easterSunday, -2);
  const easterMonday = addDays(easterSunday, 1);
  const newYearsDay = applySubstituteHoliday(new Date(Date.UTC(year, 0, 1)));
  const earlyMay = nthWeekdayOfMonth(year, 4, 1, 1);
  const springBank = lastWeekdayOfMonth(year, 4, 1);
  const summerBank = lastWeekdayOfMonth(year, 7, 1);

  let christmasDayObserved = new Date(Date.UTC(year, 11, 25));
  let boxingDayObserved = new Date(Date.UTC(year, 11, 26));
  const christmasWeekday = christmasDayObserved.getUTCDay();
  const boxingWeekday = boxingDayObserved.getUTCDay();

  if (christmasWeekday === 6) {
    christmasDayObserved = new Date(Date.UTC(year, 11, 27));
    boxingDayObserved = new Date(Date.UTC(year, 11, 28));
  } else if (christmasWeekday === 0) {
    christmasDayObserved = new Date(Date.UTC(year, 11, 27));
    boxingDayObserved = new Date(Date.UTC(year, 11, 26));
  } else if (boxingWeekday === 6 || boxingWeekday === 0) {
    boxingDayObserved = new Date(Date.UTC(year, 11, 28));
  }

  return [
    { date: toIsoDate(newYearsDay), label: "New Year's Day" },
    { date: toIsoDate(goodFriday), label: "Good Friday" },
    { date: toIsoDate(easterMonday), label: "Easter Monday" },
    { date: toIsoDate(earlyMay), label: "Early May Bank Holiday" },
    { date: toIsoDate(springBank), label: "Spring Bank Holiday" },
    { date: toIsoDate(summerBank), label: "Summer Bank Holiday" },
    { date: toIsoDate(christmasDayObserved), label: "Christmas Day" },
    { date: toIsoDate(boxingDayObserved), label: "Boxing Day" }
  ];
}

function buildHolidayYearRows(holidays, yearStart, holidayEvents = []) {
  const startMonth = new Date(Date.UTC(yearStart, 1, 1));
  const holidayMap = new Map();
  const eventMap = new Map();
  const years = new Set();
  const rows = [];

  for (let offset = 0; offset < 12; offset += 1) {
    const monthDate = addMonths(startMonth, offset);
    rows.push(monthDate);
    years.add(monthDate.getUTCFullYear());
  }

  const bankHolidayMap = new Map();
  years.forEach((year) => {
    getUkBankHolidays(year).forEach((holiday) => bankHolidayMap.set(holiday.date, holiday.label));
  });

  holidays.forEach((holiday) => {
    const bucket = holidayMap.get(holiday.date) || [];
    bucket.push(holiday);
    holidayMap.set(holiday.date, bucket);
  });

  holidayEvents.forEach((event) => {
    const bucket = eventMap.get(event.date) || [];
    bucket.push(event);
    eventMap.set(event.date, bucket);
  });

  return rows.map((monthDate) => {
    const year = monthDate.getUTCFullYear();
    const month = monthDate.getUTCMonth();
    const monthLabel = monthDate.toLocaleString("en-GB", { month: "short", year: "2-digit", timeZone: "UTC" });
    const days = Array.from({ length: 31 }, (_, index) => {
      const dayNumber = index + 1;
      const candidate = new Date(Date.UTC(year, month, dayNumber));
      const inMonth = candidate.getUTCMonth() === month;
      const isoDate = inMonth ? toIsoDate(candidate) : "";
      const weekday = candidate.getUTCDay();
      return {
        key: `${year}-${month + 1}-${dayNumber}`,
        dayNumber,
          isoDate,
          inMonth,
          weekend: inMonth && (weekday === 0 || weekday === 6),
          bankHoliday: inMonth ? bankHolidayMap.get(isoDate) || "" : "",
          holidays: inMonth ? (holidayMap.get(isoDate) || []) : [],
          events: inMonth ? (eventMap.get(isoDate) || []) : []
        };
      });

    return {
      id: `${year}-${String(month + 1).padStart(2, "0")}`,
      label: monthLabel,
      days
    };
  });
}

function createMessage(text, tone = "info") {
  return { text, tone, id: `${Date.now()}-${Math.random()}` };
}

function buildJobPhotoUrl(jobId, photoId) {
  return `/api/jobs/${encodeURIComponent(jobId)}/photos/${encodeURIComponent(photoId)}`;
}

function formatDateTime(value) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value);
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(parsed);
}

function formatJobDate(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return String(value || "");
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: "Europe/London"
  }).format(parsed);
}

function formatHolidayRequestDateRange(startDate, endDate) {
  const start = formatJobDate(startDate);
  const end = formatJobDate(endDate || startDate);
  if (!start) return "";
  return end && end !== start ? `${start} to ${end}` : start;
}

function getRamsJobTitle(job) {
  return [job?.orderReference, job?.customerName, job?.description].filter(Boolean).join(" - ") || "Untitled job";
}

function getRamsJobAddress(job) {
  return String(job?.address || "").trim() || "Site address to be confirmed";
}

function getRamsContact(job) {
  return [job?.contact, job?.number].filter(Boolean).join(" - ") || "Site contact to be confirmed";
}

function getInstallerNamesForRams(job) {
  const names = Array.isArray(job?.installers)
    ? job.installers
      .filter((entry) => entry && entry !== "Custom")
      .map((entry) => getHolidayStaffEntry(entry)?.fullName || entry)
    : [];
  if (job?.installers?.includes?.("Custom") && job?.customInstaller) names.push(job.customInstaller);
  return names.length ? names.join(", ") : "To be allocated";
}

function formatRamsCreatedDate(value) {
  const parsed = value ? new Date(value) : new Date();
  if (Number.isNaN(parsed.getTime())) return formatJobDate(getLocalTodayIso());
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: "Europe/London"
  }).format(parsed);
}

function normalizeRamsQuestions(questions) {
  return {
    ...RAMS_DEFAULT_QUESTIONS,
    ...questions,
    tools: Array.isArray(questions?.tools) && questions.tools.length ? questions.tools : RAMS_DEFAULT_QUESTIONS.tools
  };
}

function toRamsScore(value, fallback = 1) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.min(5, Math.round(parsed)));
}

function calculateRamsRisk(likelihood, consequence) {
  return toRamsScore(likelihood, 0) * toRamsScore(consequence, 0);
}

function getRamsRiskBand(rating) {
  const value = Number(rating) || 0;
  if (value >= 16) return { code: "H", label: "High risk", className: "risk-high" };
  if (value >= 11) return { code: "M", label: "Medium risk", className: "risk-medium" };
  if (value >= 5) return { code: "L", label: "Low risk / monitor", className: "risk-low" };
  return { code: "N", label: "No further action required", className: "risk-none" };
}

function normalizeRamsCard(card = {}, fallback = {}) {
  const merged = { ...fallback, ...card };
  const content = Array.isArray(merged.content)
    ? merged.content.map((line) => String(line)).filter(Boolean)
    : String(merged.content || merged.controlMeasure || "").split("\n").map((line) => line.trim()).filter(Boolean);
  const type = String(merged.type || fallback.type || "Risk");
  const isRisk = type !== "Method";
  const normalized = {
    ...merged,
    title: String(merged.title || fallback.title || "RAMS card"),
    type,
    trigger: String(merged.trigger || fallback.trigger || "Custom"),
    content: content.length ? content : ["Write the control measure or method step here."]
  };
  if (!isRisk) return normalized;
  const responsibility = String(merged.responsibility || fallback.responsibility || "Installers from selected job");
  return {
    ...normalized,
    whoAtRisk: String(merged.whoAtRisk || fallback.whoAtRisk || "Employees\nThird parties"),
    responsibility: responsibility === "Matt Carroll" ? "Installers from selected job" : responsibility,
    controlMeasure: String(merged.controlMeasure || fallback.controlMeasure || normalized.content.join("\n")),
    initialLikelihood: toRamsScore(merged.initialLikelihood ?? fallback.initialLikelihood ?? getRamsLcr(merged, "initial").likelihood, 2),
    initialConsequence: toRamsScore(merged.initialConsequence ?? fallback.initialConsequence ?? getRamsLcr(merged, "initial").consequence, 3),
    residualLikelihood: toRamsScore(merged.residualLikelihood ?? fallback.residualLikelihood ?? getRamsLcr(merged, "residual").likelihood, 1),
    residualConsequence: toRamsScore(merged.residualConsequence ?? fallback.residualConsequence ?? getRamsLcr(merged, "residual").consequence, 1)
  };
}

function normalizeRamsLogic(logic = {}) {
  const defaultGroups = RAMS_DEFAULT_LOGIC.optionGroups;
  const incomingGroups = Array.isArray(logic.optionGroups) ? logic.optionGroups : defaultGroups;
  const defaultCards = {
    ...RAMS_DEFAULT_LOGIC.cards,
    ...RAMS_STANDARD_RISK_CARDS
  };
  const incomingCards = logic.cards && typeof logic.cards === "object" ? logic.cards : {};
  const cardEntries = Object.entries({ ...defaultCards, ...incomingCards }).map(([cardId, card]) => [
    cardId,
    normalizeRamsCard(card, defaultCards[cardId])
  ]);
  return {
    optionGroups: incomingGroups.map((group, groupIndex) => {
      const fallback = defaultGroups[groupIndex] || {};
      return {
        key: String(group.key || fallback.key || `group-${groupIndex}`),
        label: String(group.label || fallback.label || "Question"),
        input: ["buttons", "select", "checkboxes"].includes(group.input) ? group.input : fallback.input || "buttons",
        multi: Boolean(group.multi ?? fallback.multi),
        options: (Array.isArray(group.options) ? group.options : []).map((option, optionIndex) => ({
          value: String(option.value || `option-${optionIndex}`),
          label: String(option.label || option.value || "Option"),
          cardIds: Array.isArray(option.cardIds) ? option.cardIds.map(String).filter(Boolean) : []
        }))
      };
    }),
    cards: Object.fromEntries(cardEntries),
    baseCardIds: Array.isArray(logic.baseCardIds)
      ? logic.baseCardIds.map(String).filter(Boolean)
      : RAMS_BASE_CARD_IDS
  };
}

function getStoredRamsLogic() {
  if (typeof window === "undefined") return normalizeRamsLogic(RAMS_DEFAULT_LOGIC);
  try {
    const stored = window.localStorage.getItem(RAMS_LOGIC_STORAGE_KEY);
    if (!stored) return normalizeRamsLogic(RAMS_DEFAULT_LOGIC);
    return normalizeRamsLogic(JSON.parse(stored));
  } catch (error) {
    console.error(error);
    return normalizeRamsLogic(RAMS_DEFAULT_LOGIC);
  }
}

function saveRamsLogic(logic) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(RAMS_LOGIC_STORAGE_KEY, JSON.stringify(normalizeRamsLogic(logic)));
}

function getRamsCardIdsForQuestions(questions, logic = RAMS_DEFAULT_LOGIC) {
  const normalized = normalizeRamsQuestions(questions);
  const activeLogic = normalizeRamsLogic(logic);
  const selected = new Set(activeLogic.baseCardIds);

  activeLogic.optionGroups.forEach((group) => {
    const answer = normalized[group.key];
    const selectedValues = group.multi
      ? Array.isArray(answer) ? answer.map(String) : []
      : [String(answer || "")];
    group.options.forEach((option) => {
      if (!selectedValues.includes(String(option.value))) return;
      option.cardIds.forEach((cardId) => selected.add(cardId));
    });
  });

  return Object.keys(activeLogic.cards).filter((cardId) => selected.has(cardId));
}

function buildRamsReference(job, questions) {
  const reference = String(job?.orderReference || job?.id || "RAMS").trim();
  const date = String(job?.date || getLocalTodayIso()).replaceAll("-", "");
  const activity = String(questions?.activity || "works").toUpperCase();
  return `${reference}-${date}-${activity}`;
}

function getRamsLcr(card, phase = "initial") {
  const explicitLikelihood = phase === "initial" ? card?.initialLikelihood : card?.residualLikelihood;
  const explicitConsequence = phase === "initial" ? card?.initialConsequence : card?.residualConsequence;
  if (explicitLikelihood !== undefined || explicitConsequence !== undefined) {
    const likelihood = toRamsScore(explicitLikelihood, phase === "initial" ? 2 : 1);
    const consequence = toRamsScore(explicitConsequence, phase === "initial" ? 3 : 1);
    return {
      likelihood,
      consequence,
      rating: calculateRamsRisk(likelihood, consequence),
      label: getRamsRiskBand(calculateRamsRisk(likelihood, consequence)).label
    };
  }
  const fallback = phase === "initial"
    ? { low: [1, 2, 2], medium: [2, 3, 6], high: [3, 4, 12] }
    : { low: [1, 1, 1], medium: [1, 2, 2], high: [2, 2, 4] };
  const riskLabel = String(phase === "initial" ? card?.initialRisk : card?.residualRisk || "Low").toLowerCase();
  const [likelihood, consequence, rating] = fallback[riskLabel] || fallback.low;
  return {
    likelihood,
    consequence,
    rating,
    label: phase === "initial" ? card?.initialRisk || "Low" : card?.residualRisk || "Low"
  };
}

function getRamsCardTypeRank(type = "") {
  const normalized = String(type || "").toLowerCase();
  if (normalized === "method") return 1;
  if (normalized === "risk") return 2;
  if (normalized === "coshh") return 3;
  return 4;
}

function sortRamsCardEntries(entries) {
  return [...entries].sort(([, left], [, right]) => {
    const rankDifference = getRamsCardTypeRank(left?.type) - getRamsCardTypeRank(right?.type);
    if (rankDifference) return rankDifference;
    return String(left?.title || "").localeCompare(String(right?.title || ""));
  });
}

function formatNotificationDate(value) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(parsed);
}

function toMonthIdFromIso(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return "";
  return `${parsed.getUTCFullYear()}-${String(parsed.getUTCMonth() + 1).padStart(2, "0")}`;
}

function shiftMonthId(value, offset) {
  const [year, month] = String(value || "").split("-").map(Number);
  if (!year || !month) return toMonthIdFromIso(getLocalTodayIso());
  const next = new Date(Date.UTC(year, month - 1 + offset, 1));
  return `${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, "0")}`;
}

function getAttendanceDisplayClass(cell) {
  const label = String(cell?.displayLabel || "").trim().toLowerCase();
  if (label.includes("unpaid") || label.includes("absence") || label.includes("absent")) return "is-unpaid";
  if (label.includes("birthday")) return "is-birthday";
  if (label.includes("bank holiday")) return "is-bank-holiday";
  if (label.includes("weekend")) return "is-weekend";
  if (label === "off") return "is-weekend";
  if (label.includes("holiday")) return "is-holiday";
  return "";
}

function formatNotificationMessage(message) {
  const raw = String(message || "").trim();
  if (!raw) return "";

  return raw
    .replace(/\b(\d{4})-(\d{2})-(\d{2})\b/g, (_, year, month, day) => `${day}/${month}/${String(year).slice(-2)}`)
    .replace(/\s+/g, " ")
    .trim();
}

function getNextWeekdayIsoDate(isoDate) {
  const parsed = parseIsoDate(isoDate);
  if (!parsed) return "";
  let cursor = addDays(parsed, 1);
  while (cursor.getUTCDay() === 0 || cursor.getUTCDay() === 6) {
    cursor = addDays(cursor, 1);
  }
  return toIsoDate(cursor);
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

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read the selected photo."));
    reader.readAsDataURL(file);
  });
}

function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Could not load the selected photo."));
    image.src = dataUrl;
  });
}

async function compressPhotoForUpload(file) {
  const originalDataUrl = await readFileAsDataUrl(file);
  const image = await loadImageFromDataUrl(originalDataUrl);
  const maxDimension = 1600;
  const scale = Math.min(1, maxDimension / Math.max(image.width || 1, image.height || 1));
  const width = Math.max(1, Math.round((image.width || 1) * scale));
  const height = Math.max(1, Math.round((image.height || 1) * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("Could not prepare the selected photo.");
  }
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob((nextBlob) => {
      if (nextBlob) resolve(nextBlob);
      else reject(new Error("Could not compress the selected photo."));
    }, "image/jpeg", 0.72);
  });

  const dataUrl = await readFileAsDataUrl(blob);
  const baseName = String(file.name || "job-photo").replace(/\.[^.]+$/, "") || "job-photo";
  return {
    fileName: `${baseName}.jpg`,
    dataUrl,
    width,
    height,
    size: blob.size
  };
}

function toInitials(name) {
  return String(name || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0])
    .join("");
}

function renderJobCardContent({
  job,
  isCondensed = false,
  isClientMode,
  draggingJobId,
  getJobTypeMeta,
  getJobTypeLabel,
  getInstallerDisplayList,
  getInstallerMeta,
  editJob,
  handleDelete,
  setActiveClientJob,
  buildDragPreview,
  getTransparentDragImage,
  clearDragPreview,
  dragPreviewRef,
  dragPositionRef,
  setDraggingJobId,
  duplicatingJobId,
  setDuplicatingJobId,
  setDropDate
}) {
  const meta = getJobTypeMeta(job.jobType);
  const installerLabels = getInstallerDisplayList(job);
  const savedRamsDocuments = Array.isArray(job.ramsDocuments) ? job.ramsDocuments : [];
  const latestRams = savedRamsDocuments[0] || null;

  return (
      <div
        key={job.id}
        className={`job-card ${meta.colorClass}-card ${job.isPlaceholder ? "is-placeholder" : ""} ${job.isCompleted ? "is-complete" : ""} ${job.isSnagging ? "is-snagging" : ""} ${isCondensed ? "is-condensed" : ""} ${draggingJobId === job.id ? "is-dragging" : ""}`}
        draggable={!isClientMode}
      onDragStart={(event) => {
        if (isClientMode) return;
        event.dataTransfer.setData("text/plain", job.id);
        event.dataTransfer.effectAllowed = "move";
        const preview = buildDragPreview(event.currentTarget);
        dragPreviewRef.current = preview;
        dragPositionRef.current = { x: event.clientX, y: event.clientY };
        preview.style.left = `${event.clientX + 18}px`;
        preview.style.top = `${event.clientY + 18}px`;
        preview.style.transform = "rotate(-2deg) translateY(0)";
        event.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0);
        setDraggingJobId(job.id);
      }}
      onDragEnd={() => {
        if (isClientMode) return;
        setDraggingJobId("");
        setDropDate("");
        clearDragPreview();
      }}
      onClick={() => {
        if (isClientMode) {
          setActiveClientJob(job);
        } else {
          editJob(job);
        }
      }}
    >
        <div className="job-card-top">
          <div className="job-title-wrap">
            <strong className="job-title-line">
              {job.orderReference ? <span className="job-ref-inline">{job.orderReference}</span> : null}
              <span className="job-customer-inline">{job.customerName}</span>
            </strong>
            <p>{job.description || "No description"}</p>
          </div>
        <div className="job-title-meta">
          {job.isPlaceholder ? <span className="placeholder-status-pill">Placeholder</span> : null}
          {job.isSnagging ? <span className="job-snagging-pill">Snagging</span> : null}
          {job.isCompleted ? <span className="job-complete-pill">Complete</span> : null}
          {latestRams ? (
            <button
              className="job-rams-pill"
              type="button"
              onClick={(event) => {
                event.stopPropagation();
                window.location.assign(`/rams?jobId=${encodeURIComponent(job.id)}&ramsId=${encodeURIComponent(latestRams.id)}`);
              }}
            >
              RAMS saved
            </button>
          ) : null}
          {Array.isArray(job.photos) && job.photos.length ? <span className="job-photo-pill">{job.photos.length} photo{job.photos.length === 1 ? "" : "s"}</span> : null}
          {installerLabels.length ? (
            <div className="job-title-installers">
              {installerLabels.map((installer) => {
                const metaInstaller = getInstallerMeta(installer);
                return (
                  <span key={`title-${job.id}-${installer}`} className={`installer-badge title-inline ${metaInstaller.colorClass}`}>
                    {installer}
                  </span>
                );
              })}
            </div>
          ) : null}
            <span className={`job-tag ${meta.colorClass}`}>{getJobTypeLabel(job)}</span>
          </div>
        </div>
        {!isCondensed ? (
          <>
            <div className="job-meta-grid">
              <p><b>Address:</b> {job.address || "-"}</p>
              <p><b>Contact:</b> {job.contact || "-"}</p>
              <p><b>Number:</b> {job.number || "-"}</p>
            </div>
            <p className="job-notes compact"><b>Notes:</b> {job.notes || ""}</p>
            <div className="job-actions">
              {!isClientMode ? (
                <>
                  <button className="text-button" type="button" onClick={(event) => { event.stopPropagation(); editJob(job); }}>
                    Edit
                  </button>
                  {latestRams ? (
                    <button
                      className="text-button"
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        window.location.assign(`/rams?jobId=${encodeURIComponent(job.id)}&ramsId=${encodeURIComponent(latestRams.id)}`);
                      }}
                    >
                      Open RAMS
                    </button>
                  ) : null}
                  <button className="text-button danger" type="button" onClick={(event) => { event.stopPropagation(); handleDelete(job.id); }}>
                    Delete
                  </button>
                  <button
                    type="button"
                    className="card-duplicate-handle"
                    draggable
                    onDragStart={(event) => {
                      event.stopPropagation();
                      event.dataTransfer.setData("job-copy", job.id);
                      event.dataTransfer.effectAllowed = "copy";
                      const preview = buildDragPreview(event.currentTarget.closest(".job-card"));
                      dragPreviewRef.current = preview;
                      dragPositionRef.current = { x: event.clientX, y: event.clientY };
                      preview.style.left = `${event.clientX + 18}px`;
                      preview.style.top = `${event.clientY + 18}px`;
                      event.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0);
                      setDuplicatingJobId(job.id);
                    }}
                    onDragEnd={() => {
                      setDuplicatingJobId("");
                      setDropDate("");
                      clearDragPreview();
                    }}
                    title="Drag to copy"
                  >
                    Drag to Copy
                  </button>
                </>
              ) : null}
            </div>
          </>
        ) : null}
      </div>
    );
  }

function getPermissionForApp(user, key) {
  const fallback =
    key === "board"
      ? user?.role === "host"
        ? "admin"
        : "user"
      : key === "vanEstimator"
        ? "none"
      : user?.role === "host"
        ? "admin"
        : "none";
  const value = String(user?.permissions?.[key] || "").trim().toLowerCase();
  return ["admin", "user", "none"].includes(value) ? value : fallback;
}

function canAccessBoard(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "board") !== "none";
}

function canEditBoard(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "board") === "admin";
}

function canAccessInstaller(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "installer") !== "none";
}

function canEditInstaller(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "installer") === "admin";
}

function canAccessHolidays(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "holidays") !== "none";
}

function canEditHolidays(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "holidays") === "admin";
}

function canAccessAttendance(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "attendance") !== "none";
}

function canEditAttendance(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "attendance") === "admin";
}

function canAccessMileage(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "mileage") !== "none";
}

function canEditMileage(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "mileage") === "admin";
}

function canAccessVanEstimator(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "vanEstimator") !== "none";
}

function canAccessRams(user) {
  return canEditBoard(user);
}

function usesHostShell(user) {
  return Boolean(
    user &&
      (canAccessInstaller(user) || canEditBoard(user) || canAccessHolidays(user) || canEditAttendance(user) || canAccessMileage(user) || canAccessVanEstimator(user) || canAccessRams(user) || user.canManagePermissions)
  );
}

function getHomePathForUser(user) {
  return usesHostShell(user) ? "/" : "/client";
}

function getBoardPathForUser(user) {
  return canEditBoard(user) ? "/board" : "/client/board";
}

function HomeIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4 11.5 12 5l8 6.5v7a1.5 1.5 0 0 1-1.5 1.5h-4.25V14h-4.5v6H5.5A1.5 1.5 0 0 1 4 18.5z" />
    </svg>
  );
}

function BoardIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M5 5h14a1 1 0 0 1 1 1v12.5A1.5 1.5 0 0 1 18.5 20h-13A1.5 1.5 0 0 1 4 18.5V6a1 1 0 0 1 1-1Zm1.5 3v9h11V8Zm0-1.5h11V6.5h-11Z" />
    </svg>
  );
}

function HolidayIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 3h2v2h6V3h2v2h1.5A1.5 1.5 0 0 1 20 6.5v12A1.5 1.5 0 0 1 18.5 20h-13A1.5 1.5 0 0 1 4 18.5v-12A1.5 1.5 0 0 1 5.5 5H7Zm11 6.5h-12v8.5h12Zm-9.5 2h3v3h-3Z" />
    </svg>
  );
}

function AttendanceIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 4h10a2 2 0 0 1 2 2v11a3 3 0 0 1-3 3H8a3 3 0 0 1-3-3V6a2 2 0 0 1 2-2Zm0 2v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V6Zm2 2h6v2H9Zm0 4h4v2H9Z" />
    </svg>
  );
}

function MileageIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M6.5 5a3.5 3.5 0 0 0-3.5 3.5c0 2.5 3.5 6.5 3.5 6.5S10 11 10 8.5A3.5 3.5 0 0 0 6.5 5Zm0 4.7a1.2 1.2 0 1 1 0-2.4 1.2 1.2 0 0 1 0 2.4ZM17.5 3A3.5 3.5 0 0 0 14 6.5c0 2.5 3.5 6.5 3.5 6.5S21 9 21 6.5A3.5 3.5 0 0 0 17.5 3Zm0 4.7a1.2 1.2 0 1 1 0-2.4 1.2 1.2 0 0 1 0 2.4ZM7 18h11a1 1 0 1 1 0 2H7a4 4 0 0 1-4-4h2a2 2 0 0 0 2 2Zm10-2a4 4 0 0 1-4-4h2a2 2 0 0 0 2 2h1v2Z" />
    </svg>
  );
}

function VanEstimatorIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4.5 8.5 6.2 5h8.9a2 2 0 0 1 1.7.94l1.58 2.56H20a1 1 0 0 1 1 1v5.25a1.25 1.25 0 0 1-1.25 1.25h-.85a2.5 2.5 0 0 1-4.8 0H9.9a2.5 2.5 0 0 1-4.8 0h-.85A1.25 1.25 0 0 1 3 14.75V10a1.5 1.5 0 0 1 1.5-1.5Zm3-1.5-.75 1.5h4.75V7Zm6 0v1.5h2.55L15.11 7Zm-6 10a1 1 0 1 0 0-2 1 1 0 0 0 0 2Zm9 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2Z" />
    </svg>
  );
}

function RamsIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 3h8.25L20 7.75V19a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Zm7 1.8V9h4.2ZM8 11h8v1.6H8Zm0 3.2h8v1.6H8Zm0 3.2h5v1.6H8Z" />
    </svg>
  );
}

function InstallerIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M6 4h12a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2h-4l-2 2-2-2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2Zm1.5 4v2h9V8Zm0 4v2h6v-2Z" />
    </svg>
  );
}

function NotificationIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 4a4 4 0 0 1 4 4v1.35c0 .83.27 1.63.78 2.29l1.28 1.69c.68.9.04 2.17-1.09 2.17H6.03c-1.13 0-1.77-1.27-1.09-2.17l1.28-1.69A3.75 3.75 0 0 0 7 9.35V8a5 5 0 0 1 5-5Zm0 16a2.75 2.75 0 0 0 2.58-1.8h-5.16A2.75 2.75 0 0 0 12 20Z" />
    </svg>
  );
}

function LogoutIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M10 4H6.5A1.5 1.5 0 0 0 5 5.5v13A1.5 1.5 0 0 0 6.5 20H10v-2H7V6h3Zm6.3 3.8-1.4 1.4 1.8 1.8H10v2h6.7l-1.8 1.8 1.4 1.4L20.5 12z" />
    </svg>
  );
}

function BrandLogoIcon() {
  return <img src="/branding/signs-express-logo.svg" alt="Signs Express" className="host-nav-brand-logo" />;
}

function getNotificationCategory(notification) {
  const title = String(notification?.title || "").toLowerCase();
  const message = String(notification?.message || "").toLowerCase();
  const link = String(notification?.link || "").toLowerCase();

  if (link.includes("/holidays") || title.includes("holiday") || message.includes("holiday")) {
    return { label: "Holiday", icon: HolidayIcon, className: "notification-type-holiday" };
  }
  if (link.includes("/attendance") || title.includes("clock") || message.includes("clock")) {
    return { label: "Attendance", icon: AttendanceIcon, className: "notification-type-update" };
  }
  if (link.includes("/mileage") || title.includes("mileage") || message.includes("miles")) {
    return { label: "Mileage", icon: MileageIcon, className: "notification-type-update" };
  }
  if (link.includes("/board") || title.includes("job") || message.includes("job")) {
    return { label: "Board", icon: BoardIcon, className: "notification-type-board" };
  }
  if (title.includes("message") || message.includes("message")) {
    return { label: "Message", icon: NotificationIcon, className: "notification-type-message" };
  }
  return { label: "Update", icon: NotificationIcon, className: "notification-type-update" };
}

function buildBoardUrl(startIso = "", endIso = "") {
  if (startIso && endIso) {
    return `/api/board?start=${encodeURIComponent(startIso)}&end=${encodeURIComponent(endIso)}`;
  }
  return "/api/board";
}

function MainNavBar({
  currentUser,
  active = "home",
  onLogout,
  notifications = []
}) {
  function goTo(path) {
    window.location.assign(path);
  }

  const boardAllowed = canAccessBoard(currentUser);
  const attendanceAllowed = canAccessAttendance(currentUser);
  const holidaysAllowed = canAccessHolidays(currentUser);
  const mileageAllowed = canAccessMileage(currentUser);
  const vanEstimatorAllowed = canAccessVanEstimator(currentUser);
  const ramsAllowed = canAccessRams(currentUser);
  const installerAllowed = canAccessInstaller(currentUser);
  const homePath = getHomePathForUser(currentUser);
  const boardPath = getBoardPathForUser(currentUser);
  const attendancePath = "/attendance";
  const holidaysPath = "/holidays";
  const mileagePath = "/mileage";
  const vanEstimatorPath = "/van-estimator";
  const ramsPath = "/rams";
  const installerPath = "/installer";
  const notificationsPath = "/notifications";
  const unreadNotifications = notifications.filter((entry) => !entry.read);

  return (
    <header className="host-nav-shell">
      <nav className="host-nav">
        <div className="host-nav-inner">
          <button type="button" className="host-nav-brand" onClick={() => goTo(homePath)} aria-label="Go to home">
            <BrandLogoIcon />
          </button>
          <div className="host-nav-links">
            <button
              type="button"
              className={`host-nav-link ${active === "home" ? "active" : ""}`}
              onClick={() => goTo(homePath)}
            >
              <span className="host-nav-link-label">Home</span>
            </button>
            {boardAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "board" ? "active" : ""}`}
                onClick={() => goTo(boardPath)}
              >
                <span className="host-nav-link-label">Installation Board</span>
              </button>
            ) : null}
            {attendanceAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "attendance" ? "active" : ""}`}
                onClick={() => goTo(attendancePath)}
              >
                <span className="host-nav-link-label">Attendance</span>
              </button>
            ) : null}
            {holidaysAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "holidays" ? "active" : ""}`}
                onClick={() => goTo(holidaysPath)}
              >
                <span className="host-nav-link-label">Holidays</span>
              </button>
            ) : null}
            {mileageAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "mileage" ? "active" : ""}`}
                onClick={() => goTo(mileagePath)}
              >
                <span className="host-nav-link-label">Mileage</span>
              </button>
            ) : null}
            <button
              type="button"
              className={`host-nav-link ${active === "van-estimator" ? "active" : ""} ${vanEstimatorAllowed ? "" : "disabled"}`}
              disabled={!vanEstimatorAllowed}
              onClick={() => goTo(vanEstimatorPath)}
              title={vanEstimatorAllowed ? "Open Vehicle Pricing Calculator" : "Vehicle Pricing Calculator is inactive"}
            >
              <span className="host-nav-link-label">Vehicle Pricing Calculator</span>
            </button>
            {ramsAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "rams" ? "active" : ""}`}
                onClick={() => goTo(ramsPath)}
              >
                <span className="host-nav-link-label">RAMS</span>
              </button>
            ) : null}
            {installerAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "installer" ? "active" : ""}`}
                onClick={() => goTo(installerPath)}
              >
                <span className="host-nav-link-label">Subcontractor Directory</span>
              </button>
            ) : null}
            <button
              type="button"
              className={`host-nav-link ${active === "notifications" ? "active" : ""}`}
              onClick={() => goTo(notificationsPath)}
            >
              <span className="host-nav-link-label">
                Notifications
                {unreadNotifications.length ? <span className="host-nav-badge inline">{unreadNotifications.length}</span> : null}
              </span>
            </button>
          </div>
          <div className="host-nav-meta">
            <span className="host-nav-user">Logged in as <strong>{currentUser.displayName}</strong></span>
            <button className="host-nav-logout" type="button" onClick={onLogout}>
              <span className="host-nav-link-label">Log out</span>
            </button>
          </div>
        </div>
      </nav>
    </header>
  );
}

function PermissionsPanel({
  currentUser,
  users,
  savingKey,
  onChangePermission,
  onUpdateAttendanceProfile,
  onCreateUser,
  onResetPassword,
  onDeleteUser
}) {
  const [createForm, setCreateForm] = useState({ displayName: "", role: "client", password: "" });
  const [passwordDrafts, setPasswordDrafts] = useState({});
  const [attendanceDrafts, setAttendanceDrafts] = useState({});
  const visibleUsers = [...users].sort((left, right) => left.displayName.localeCompare(right.displayName));

  useEffect(() => {
    setAttendanceDrafts(
      Object.fromEntries(
        users.map((user) => [user.id, normalizeAttendanceDraft(user.attendanceProfile)])
      )
    );
  }, [users]);

  function updateAttendanceDraft(userId, updater) {
    setAttendanceDrafts((current) => {
      const existing = current[userId] || normalizeAttendanceDraft(null);
      const nextValue = typeof updater === "function" ? updater(existing) : updater;
      return {
        ...current,
        [userId]: normalizeAttendanceDraft(nextValue)
      };
    });
  }

  return (
      <section className="panel permissions-panel">
        <div className="permissions-head">
          <h3>User portal</h3>
          <p>Manage access, passwords, contracted hours and attendance rules in one place.</p>
        </div>
        <div className="permissions-admin-tools">
        <input
          type="text"
          placeholder="Full name"
          value={createForm.displayName}
          onChange={(event) => setCreateForm((current) => ({ ...current, displayName: event.target.value }))}
        />
        <select
          value={createForm.role}
          onChange={(event) => setCreateForm((current) => ({ ...current, role: event.target.value }))}
        >
          <option value="client">Client</option>
          <option value="host">Host</option>
        </select>
        <input
          type="password"
          placeholder="Temporary password"
          value={createForm.password}
          onChange={(event) => setCreateForm((current) => ({ ...current, password: event.target.value }))}
        />
        <button
          className="primary-button"
          type="button"
          onClick={async () => {
            await onCreateUser(createForm);
            setCreateForm({ displayName: "", role: "client", password: "" });
          }}
          disabled={!createForm.displayName.trim() || !createForm.password}
        >
          Add user
        </button>
      </div>
      <div className="permissions-grid">
          {visibleUsers.map((user) => {
            const isSelf = user.id === currentUser.id;
            const boardPermission = getPermissionForApp(user, "board");
            const holidaysPermission = getPermissionForApp(user, "holidays");
            const installerPermission = getPermissionForApp(user, "installer");
            const attendancePermission = getPermissionForApp(user, "attendance");
            const mileagePermission = getPermissionForApp(user, "mileage");
            const vanEstimatorPermission = getPermissionForApp(user, "vanEstimator");
            const attendanceProfile = normalizeAttendanceDraft(user.attendanceProfile);
            const attendanceDraft = attendanceDrafts[user.id] || attendanceProfile;
            const attendanceMode = String(attendanceDraft.mode || "required");
            const contractedHours = attendanceDraft.contractedHours || {};
            const exemptFromClocking = attendanceMode === "exempt";
            const fixedHoursMode = attendanceMode === "fixed";
            const attendanceChanged = JSON.stringify(attendanceDraft) !== JSON.stringify(attendanceProfile);

            return (
              <article key={user.id} className="permissions-user-card">
              <div className="permissions-user-head">
                <div className="permissions-user-identity">
                  <strong>{user.displayName}</strong>
                  <span className="permissions-user-meta">{user.role === "host" ? "Host" : "Client"}</span>
                  </div>
                  {isSelf ? <span className="permissions-owner-pill">Owner</span> : null}
                </div>

                <div className="permissions-user-body">
                  <div className="permissions-main-grid">
                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Installation Board</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-board-${option.value}`}
                            type="button"
                            className={`permission-chip ${boardPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:board`}
                            onClick={() => onChangePermission(user.id, "board", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Subcontractor Directory</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-installer-${option.value}`}
                            type="button"
                            className={`permission-chip ${installerPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:installer`}
                            onClick={() => onChangePermission(user.id, "installer", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Holidays</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-holidays-${option.value}`}
                            type="button"
                            className={`permission-chip ${holidaysPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:holidays`}
                            onClick={() => onChangePermission(user.id, "holidays", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Attendance</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-attendance-${option.value}`}
                            type="button"
                            className={`permission-chip ${attendancePermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:attendance`}
                            onClick={() => onChangePermission(user.id, "attendance", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Mileage</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-mileage-${option.value}`}
                            type="button"
                            className={`permission-chip ${mileagePermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:mileage`}
                            onClick={() => onChangePermission(user.id, "mileage", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Vehicle Pricing Calculator</span>
                      <div className="permission-segment permission-toggle-segment">
                        <button
                          type="button"
                          className={`permission-chip permission-toggle-chip ${vanEstimatorPermission !== "none" ? "active active-state" : ""}`}
                          disabled={isSelf || savingKey === `${user.id}:vanEstimator`}
                          onClick={() => onChangePermission(user.id, "vanEstimator", "admin")}
                        >
                          Active
                        </button>
                        <button
                          type="button"
                          className={`permission-chip permission-toggle-chip ${vanEstimatorPermission === "none" ? "active inactive-state" : ""}`}
                          disabled={isSelf || savingKey === `${user.id}:vanEstimator`}
                          onClick={() => onChangePermission(user.id, "vanEstimator", "none")}
                        >
                          Inactive
                        </button>
                      </div>
                    </div>
                  </div>

                  <div className="permissions-attendance-settings">
                    <div className="permissions-attendance-head">
                      <div>
                        <strong>Attendance profile</strong>
                        <p>
                          {exemptFromClocking
                            ? "Removed from the attendance board."
                            : fixedHoursMode
                            ? "Uses contracted hours instead of clockings."
                            : "Uses live clock in / out times."}
                        </p>
                      </div>
                      {savingKey === `${user.id}:attendance-profile` ? (
                        <span className="permissions-saving-pill">Saving...</span>
                      ) : null}
                    </div>
                    <div className="permissions-attendance-toggles">
                      <label className="permissions-toggle">
                        <input
                          type="checkbox"
                          checked={exemptFromClocking}
                          disabled={savingKey === `${user.id}:attendance-profile`}
                          onChange={(event) => {
                            const checked = event.target.checked;
                            const nextMode = checked ? "exempt" : fixedHoursMode ? "fixed" : "required";
                            updateAttendanceDraft(user.id, (current) => ({
                              ...current,
                              mode: nextMode
                            }));
                          }}
                        />
                        <span>Exempt from clocking in / out</span>
                      </label>
                      <label className="permissions-toggle">
                        <input
                          type="checkbox"
                          checked={fixedHoursMode}
                          disabled={exemptFromClocking || savingKey === `${user.id}:attendance-profile`}
                          onChange={(event) => {
                            const checked = event.target.checked;
                            updateAttendanceDraft(user.id, (current) => ({
                              ...current,
                              mode: checked ? "fixed" : "required"
                            }));
                          }}
                        />
                        <span>No clocking in / out required</span>
                      </label>
                    </div>

                    <div className="permissions-hours-grid">
                      {ATTENDANCE_WEEKDAYS.map(([dayKey, dayLabel]) => (
                        <div key={`${user.id}-${dayKey}`} className="permissions-hours-row">
                          <span className="permissions-hours-day">{dayLabel}</span>
                          <input
                            type="text"
                            inputMode="numeric"
                            placeholder="09:00"
                            className="permissions-hours-input"
                            value={contractedHours?.[dayKey]?.in || ""}
                            disabled={Boolean(contractedHours?.[dayKey]?.off) || savingKey === `${user.id}:attendance-profile`}
                            onChange={(event) =>
                              updateAttendanceDraft(user.id, (current) => ({
                                ...current,
                                contractedHours: {
                                  ...current.contractedHours,
                                  [dayKey]: {
                                    ...(current.contractedHours?.[dayKey] || {}),
                                    in: event.target.value
                                  }
                                }
                              }))
                            }
                          />
                          <input
                            type="text"
                            inputMode="numeric"
                            placeholder="17:00"
                            className="permissions-hours-input"
                            value={contractedHours?.[dayKey]?.out || ""}
                            disabled={Boolean(contractedHours?.[dayKey]?.off) || savingKey === `${user.id}:attendance-profile`}
                            onChange={(event) =>
                              updateAttendanceDraft(user.id, (current) => ({
                                ...current,
                                contractedHours: {
                                  ...current.contractedHours,
                                  [dayKey]: {
                                    ...(current.contractedHours?.[dayKey] || {}),
                                    out: event.target.value
                                  }
                                }
                              }))
                            }
                          />
                          <label className="permissions-hours-off">
                            <input
                              type="checkbox"
                              checked={Boolean(contractedHours?.[dayKey]?.off)}
                              disabled={savingKey === `${user.id}:attendance-profile`}
                              onChange={(event) =>
                                updateAttendanceDraft(user.id, (current) => ({
                                  ...current,
                                  contractedHours: {
                                    ...current.contractedHours,
                                    [dayKey]: {
                                      ...(current.contractedHours?.[dayKey] || {}),
                                      off: event.target.checked
                                    }
                                  }
                                }))
                              }
                            />
                            <span>Off</span>
                          </label>
                        </div>
                      ))}
                    </div>

                    <div className="permissions-attendance-actions">
                      <button
                        className="ghost-button"
                        type="button"
                        disabled={!attendanceChanged || savingKey === `${user.id}:attendance-profile`}
                        onClick={() => updateAttendanceDraft(user.id, attendanceProfile)}
                      >
                        Reset
                      </button>
                      <button
                        className="primary-button"
                        type="button"
                        disabled={!attendanceChanged || savingKey === `${user.id}:attendance-profile`}
                        onClick={() => onUpdateAttendanceProfile(user.id, attendanceDraft)}
                      >
                        Save attendance settings
                      </button>
                    </div>
                  </div>
                </div>

              <div className="permissions-user-actions">
                <input
                  type="password"
                  className="permissions-password-input"
                  placeholder={user.hasPassword ? "New password" : "Set password"}
                  value={passwordDrafts[user.id] || ""}
                  onChange={(event) =>
                    setPasswordDrafts((current) => ({
                      ...current,
                      [user.id]: event.target.value
                    }))
                  }
                />
                <button
                  className="ghost-button"
                  type="button"
                  onClick={async () => {
                    await onResetPassword(user.id, passwordDrafts[user.id] || "");
                    setPasswordDrafts((current) => ({ ...current, [user.id]: "" }));
                  }}
                  disabled={!passwordDrafts[user.id]}
                >
                  Update password
                </button>
                <button
                  className="text-button danger"
                  type="button"
                  onClick={() => onDeleteUser(user)}
                  disabled={isSelf}
                >
                  Delete user
                </button>
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}

function NotificationsPage({
  currentUser,
  onLogout,
  notifications,
  onOpenNotification,
  onMarkNotificationRead,
  onMarkAllNotificationsRead
}) {
  const [activeFilter, setActiveFilter] = useState("all");
  const filterOptions = [
    { value: "all", label: "All" },
    { value: "unread", label: "Unread" },
    { value: "holiday", label: "Holidays" },
    { value: "board", label: "Jobs" },
    { value: "mileage", label: "Mileage" },
    { value: "message", label: "Messages" }
  ];
  const unreadCount = notifications.filter((entry) => !entry.read).length;
  const filteredNotifications = notifications.filter((notification) => {
    const category = getNotificationCategory(notification).label.toLowerCase();
    if (activeFilter === "all") return true;
    if (activeFilter === "unread") return !notification.read;
    return category === activeFilter;
  });

  return (
    <div className="app-shell notifications-shell">
      <div className="page notifications-page">
        <MainNavBar currentUser={currentUser} active="notifications" onLogout={onLogout} notifications={notifications} />

        <section className="panel notifications-panel">
          <div className="notifications-panel-head">
            <div>
              <h2>Notifications</h2>
              <p>{unreadCount ? `${unreadCount} unread notification${unreadCount === 1 ? "" : "s"}` : "You're all caught up."}</p>
            </div>
            {notifications.length ? (
              <button className="ghost-button" type="button" onClick={onMarkAllNotificationsRead}>
                Mark all read
              </button>
            ) : null}
          </div>

          <div className="notifications-filter-row">
            {filterOptions.map((option) => (
              <button
                key={option.value}
                type="button"
                className={`notification-filter-chip ${activeFilter === option.value ? "active" : ""}`}
                onClick={() => setActiveFilter(option.value)}
              >
                {option.label}
              </button>
            ))}
          </div>

          {filteredNotifications.length ? (
                <div className="notifications-feed">
                  {filteredNotifications.map((notification) => {
                    const category = getNotificationCategory(notification);
                    const CategoryIcon = category.icon;
                    const formattedMessage = formatNotificationMessage(notification.message);
                    const formattedTimestamp = formatNotificationDate(notification.createdAt);
                    return (
                      <article
                        key={notification.id}
                        className={`notification-feed-card ${notification.read ? "read" : "unread"}`}
                      >
                        <button
                          type="button"
                          className="notification-feed-main"
                          onClick={() => onOpenNotification(notification)}
                        >
                          <span className={`notification-feed-icon ${category.className}`}>
                            <CategoryIcon />
                          </span>
                          <div className="notification-feed-copy">
                            <div className="notification-feed-top">
                              <div className="notification-feed-title-row">
                                <strong>{notification.title}</strong>
                                {formattedTimestamp ? (
                                  <time className="notification-feed-time" dateTime={notification.createdAt}>
                                    {formattedTimestamp}
                                  </time>
                                ) : null}
                              </div>
                              <div className="notification-feed-meta-row">
                                <span className={`notification-feed-tag ${category.className}`}>{category.label}</span>
                                {!notification.read ? <span className="notification-feed-status">Unread</span> : null}
                              </div>
                            </div>
                            <p className="notification-feed-message">{formattedMessage || notification.message}</p>
                          </div>
                        </button>
                        <div className="notification-feed-actions">
                          {!notification.read ? (
                            <button
                              type="button"
                              className="text-button"
                              onClick={() => onMarkNotificationRead(notification.id)}
                            >
                              Mark read
                            </button>
                          ) : null}
                        </div>
                      </article>
                    );
                  })}
              </div>
            ) : (
            <div className="notifications-empty">No notifications in this view.</div>
          )}
        </section>
      </div>
    </div>
  );
}

function HostLandingPage({
  currentUser,
  onLogout,
  users,
  savingKey,
  onChangePermission,
  onUpdateAttendanceProfile,
  onCreateUser,
  onResetPassword,
  onDeleteUser,
  notifications
}) {
  const [permissionsOpen, setPermissionsOpen] = useState(false);

  function goTo(path) {
    window.location.assign(path);
  }

  return (
    <div className="app-shell host-landing-shell">
      <div className="page host-landing-page">
        <MainNavBar
          currentUser={currentUser}
          active="home"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel host-landing-panel">
          <div className="host-landing-actions">
            {canAccessAttendance(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/attendance")}>
                <strong>Attendance</strong>
              </button>
            ) : null}
            {canAccessHolidays(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/holidays")}>
                <strong>Holidays</strong>
              </button>
            ) : null}
            {canAccessMileage(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/mileage")}>
                <strong>Mileage</strong>
              </button>
            ) : null}
            <button
              className={`host-launch-card ${canAccessVanEstimator(currentUser) ? "" : "disabled"}`}
              type="button"
              disabled={!canAccessVanEstimator(currentUser)}
              onClick={() => goTo("/van-estimator")}
              title={canAccessVanEstimator(currentUser) ? "Open Vehicle Pricing Calculator" : "Vehicle Pricing Calculator is inactive"}
            >
              <strong>Vehicle Pricing Calculator</strong>
              {!canAccessVanEstimator(currentUser) ? <span className="host-launch-status">Inactive</span> : null}
            </button>
            {canAccessRams(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/rams")}>
                <strong>RAMS</strong>
              </button>
            ) : null}
            {canAccessInstaller(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/installer")}>
                <strong>Subcontractor Database</strong>
              </button>
            ) : null}
            {canAccessBoard(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo(getBoardPathForUser(currentUser))}>
                <strong>Installation Board</strong>
              </button>
            ) : null}
            {currentUser?.canManagePermissions ? (
              <button className="host-launch-card" type="button" onClick={() => setPermissionsOpen(true)}>
                <strong>Manage Permissions</strong>
              </button>
            ) : null}
          </div>
        </section>

      </div>

      {currentUser?.canManagePermissions && permissionsOpen ? (
        <div className="modal-backdrop" onClick={() => setPermissionsOpen(false)}>
          <div className="modal permissions-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>Permissions</h3>
                <p>Set who can use each part of the system.</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setPermissionsOpen(false)}>
                x
              </button>
            </div>
              <PermissionsPanel
                currentUser={currentUser}
                users={users}
                savingKey={savingKey}
                onChangePermission={onChangePermission}
                onUpdateAttendanceProfile={onUpdateAttendanceProfile}
                onCreateUser={onCreateUser}
                onResetPassword={onResetPassword}
                onDeleteUser={onDeleteUser}
              />
          </div>
        </div>
      ) : null}
    </div>
  );
}

function ClientLandingPage({
  currentUser,
  onLogout,
  notifications
}) {
  function goTo(path) {
    window.location.assign(path);
  }

  return (
    <div className="app-shell host-landing-shell client-landing-shell">
      <div className="page host-landing-page">
        <MainNavBar
          currentUser={currentUser}
          active="home"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel host-landing-panel">
          <div className="host-landing-actions">
            {canAccessAttendance(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/attendance")}>
                <strong>Attendance</strong>
              </button>
            ) : null}
            {canAccessHolidays(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/holidays")}>
                <strong>Holidays</strong>
              </button>
            ) : null}
            {canAccessMileage(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/mileage")}>
                <strong>Mileage</strong>
              </button>
            ) : null}
            <button
              className={`host-launch-card ${canAccessVanEstimator(currentUser) ? "" : "disabled"}`}
              type="button"
              disabled={!canAccessVanEstimator(currentUser)}
              onClick={() => goTo("/van-estimator")}
              title={canAccessVanEstimator(currentUser) ? "Open Vehicle Pricing Calculator" : "Vehicle Pricing Calculator is inactive"}
            >
              <strong>Vehicle Pricing Calculator</strong>
              {!canAccessVanEstimator(currentUser) ? <span className="host-launch-status">Inactive</span> : null}
            </button>
            {canAccessInstaller(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/installer")}>
                <strong>Subcontractor Directory</strong>
              </button>
            ) : null}
            {canAccessBoard(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo(getBoardPathForUser(currentUser))}>
                <strong>Installation Board</strong>
              </button>
            ) : null}
          </div>
        </section>

      </div>
    </div>
  );
}

function RamsLogicPage({ currentUser, onLogout, notifications }) {
  const [ramsLogicDraft, setRamsLogicDraft] = useState(() => getStoredRamsLogic());
  const [logicStatus, setLogicStatus] = useState("");
  const [selectedGroupIndex, setSelectedGroupIndex] = useState(0);
  const [selectedOptionIndex, setSelectedOptionIndex] = useState(0);
  const [selectedCardId, setSelectedCardId] = useState("");
  const [logicSection, setLogicSection] = useState("risk");

  const activeGroup = ramsLogicDraft.optionGroups[selectedGroupIndex] || ramsLogicDraft.optionGroups[0] || null;
  const activeGroupIndex = Math.max(0, ramsLogicDraft.optionGroups.indexOf(activeGroup));
  const activeOption = activeGroup?.options?.[selectedOptionIndex] || activeGroup?.options?.[0] || null;
  const activeOptionIndex = activeGroup && activeOption ? Math.max(0, activeGroup.options.indexOf(activeOption)) : 0;
  const sectionMatchesCard = (card) => logicSection === "risk" ? card?.type !== "Method" : card?.type === "Method";
  const sectionCards = sortRamsCardEntries(Object.entries(ramsLogicDraft.cards).filter(([, card]) => sectionMatchesCard(card)));
  const activeCardId = selectedCardId && ramsLogicDraft.cards[selectedCardId]
    && sectionMatchesCard(ramsLogicDraft.cards[selectedCardId])
    ? selectedCardId
    : sectionCards[0]?.[0] || "";
  const activeCard = activeCardId ? ramsLogicDraft.cards[activeCardId] : null;
  const riskBankCount = Object.values(ramsLogicDraft.cards).filter((card) => card?.type !== "Method").length;
  const methodBankCount = Object.values(ramsLogicDraft.cards).filter((card) => card?.type === "Method").length;
  const linkedTagLabels = activeCardId
    ? ramsLogicDraft.optionGroups.flatMap((group, groupIndex) =>
      group.options
        .map((option, optionIndex) => ({ group, groupIndex, option, optionIndex }))
        .filter(({ option }) => option.cardIds.includes(activeCardId))
        .map(({ group, groupIndex, option, optionIndex }) => ({
          key: `${group.key}-${option.value}-${optionIndex}`,
          label: `${group.label}: ${option.label}`,
          groupIndex,
          optionIndex
        }))
    )
    : [];

  useEffect(() => {
    if (!activeGroup) {
      setSelectedGroupIndex(0);
      setSelectedOptionIndex(0);
      return;
    }
    if (!activeOption) {
      setSelectedOptionIndex(0);
    }
  }, [activeGroup, activeOption]);

  useEffect(() => {
    if (activeCardId && activeCardId !== selectedCardId) {
      setSelectedCardId(activeCardId);
    }
  }, [activeCardId, selectedCardId]);

  function updateRamsLogicDraft(updater) {
    setRamsLogicDraft((current) => normalizeRamsLogic(typeof updater === "function" ? updater(current) : updater));
    setLogicStatus("");
  }

  function saveRamsLogicDraft() {
    const nextLogic = normalizeRamsLogic(ramsLogicDraft);
    setRamsLogicDraft(nextLogic);
    saveRamsLogic(nextLogic);
    setLogicStatus("RAMS logic saved.");
  }

  function restoreDefaultRamsLogic() {
    const nextLogic = normalizeRamsLogic(RAMS_DEFAULT_LOGIC);
    setRamsLogicDraft(nextLogic);
    saveRamsLogic(nextLogic);
    setLogicStatus("Default RAMS logic restored.");
  }

  function addRamsOption(groupIndex) {
    updateRamsLogicDraft((current) => {
      const optionGroups = current.optionGroups.map((group, index) => {
        if (index !== groupIndex) return group;
        const nextIndex = group.options.length + 1;
        return {
          ...group,
          options: [
            ...group.options,
            { value: `${group.key}-${Date.now()}`, label: `New option ${nextIndex}`, cardIds: [] }
          ]
        };
      });
      return { ...current, optionGroups };
    });
  }

  function updateRamsOption(groupIndex, optionIndex, key, value) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) => {
        if (currentGroupIndex !== groupIndex) return group;
        return {
          ...group,
          options: group.options.map((option, currentOptionIndex) =>
            currentOptionIndex === optionIndex ? { ...option, [key]: value } : option
          )
        };
      })
    }));
  }

  function removeRamsOption(groupIndex, optionIndex) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) =>
        currentGroupIndex === groupIndex
          ? { ...group, options: group.options.filter((_, currentOptionIndex) => currentOptionIndex !== optionIndex) }
          : group
      )
    }));
  }

  function toggleOptionCard(groupIndex, optionIndex, cardId) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) => {
        if (currentGroupIndex !== groupIndex) return group;
        return {
          ...group,
          options: group.options.map((option, currentOptionIndex) => {
            if (currentOptionIndex !== optionIndex) return option;
            const cardIds = option.cardIds.includes(cardId)
              ? option.cardIds.filter((entry) => entry !== cardId)
              : [...option.cardIds, cardId];
            return { ...option, cardIds };
          })
        };
      })
    }));
  }

  function setGroupCardLinks(groupIndex, cardId, shouldLink) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) => {
        if (currentGroupIndex !== groupIndex) return group;
        return {
          ...group,
          options: group.options.map((option) => {
            const hasCard = option.cardIds.includes(cardId);
            if (shouldLink && !hasCard) return { ...option, cardIds: [...option.cardIds, cardId] };
            if (!shouldLink && hasCard) return { ...option, cardIds: option.cardIds.filter((entry) => entry !== cardId) };
            return option;
          })
        };
      })
    }));
  }

  function clearCardTags(cardId) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group) => ({
        ...group,
        options: group.options.map((option) => ({
          ...option,
          cardIds: option.cardIds.filter((entry) => entry !== cardId)
        }))
      }))
    }));
  }

  function addRamsCard() {
    const cardId = `custom-${Date.now()}`;
    const isMethod = logicSection === "method";
    updateRamsLogicDraft((current) => ({
      ...current,
      cards: {
        ...current.cards,
        [cardId]: {
          title: isMethod ? "New method step" : "New risk assessment item",
          type: isMethod ? "Method" : "Risk",
          trigger: "Custom",
          initialRisk: "Medium",
          residualRisk: "Low",
          whoAtRisk: "Employees\nThird parties",
          initialLikelihood: 2,
          initialConsequence: 3,
          residualLikelihood: 1,
          residualConsequence: 1,
          responsibility: "Installers from selected job",
          controlMeasure: "Write the required control measure here.",
          content: [isMethod ? "Write the method step here." : "Write the required control measure here."]
        }
      }
    }));
    setSelectedCardId(cardId);
  }

  function updateRamsCard(cardId, key, value) {
    updateRamsLogicDraft((current) => ({
      ...current,
      cards: {
        ...current.cards,
        [cardId]: {
          ...current.cards[cardId],
          [key]: key === "content" ? String(value).split("\n").map((line) => line.trim()).filter(Boolean) : value,
          ...(key === "controlMeasure"
            ? { content: String(value).split("\n").map((line) => line.trim()).filter(Boolean) }
            : {})
        }
      }
    }));
  }

  function removeRamsCard(cardId) {
    updateRamsLogicDraft((current) => {
      const nextCards = { ...current.cards };
      delete nextCards[cardId];
      return {
        ...current,
        cards: nextCards,
        baseCardIds: current.baseCardIds.filter((entry) => entry !== cardId),
        optionGroups: current.optionGroups.map((group) => ({
          ...group,
          options: group.options.map((option) => ({
            ...option,
            cardIds: option.cardIds.filter((entry) => entry !== cardId)
          }))
        }))
      };
    });
  }

  function toggleBaseRamsCard(cardId) {
    updateRamsLogicDraft((current) => ({
      ...current,
      baseCardIds: current.baseCardIds.includes(cardId)
        ? current.baseCardIds.filter((entry) => entry !== cardId)
        : [...current.baseCardIds, cardId]
    }));
  }

  return (
    <div className="app-shell rams-shell rams-logic-shell">
      <div className="page rams-page rams-logic-page">
        <MainNavBar
          currentUser={currentUser}
          active="rams"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel rams-panel rams-logic-full-panel">
          <div className="rams-header rams-logic-full-header">
            <div>
              <span className="panel-kicker">Admin module</span>
              <h2>RAMS Logic</h2>
              <p>Manage the Risk Assessment bank and Method Statement bank, then tag each item to the job buttons that should trigger it.</p>
            </div>
            <div className="rams-logic-actions">
              <button className="ghost-button" type="button" onClick={() => window.location.assign("/rams")}>
                Back to RAMS
              </button>
              <button className="ghost-button" type="button" onClick={restoreDefaultRamsLogic}>
                Restore defaults
              </button>
              <button className="primary-button" type="button" onClick={saveRamsLogicDraft}>
                Save logic
              </button>
            </div>
          </div>
          {logicStatus ? <p className="rams-logic-status">{logicStatus}</p> : null}

          <div className="rams-logic-section-tabs" role="tablist" aria-label="RAMS logic section">
            <button
              type="button"
              className={logicSection === "risk" ? "active" : ""}
              onClick={() => {
                setLogicSection("risk");
                setSelectedCardId("");
              }}
            >
              <strong>Risk Assessment</strong>
              <span>{riskBankCount} hazards</span>
            </button>
            <button
              type="button"
              className={logicSection === "method" ? "active" : ""}
              onClick={() => {
                setLogicSection("method");
                setSelectedCardId("");
              }}
            >
              <strong>Method Statement</strong>
              <span>{methodBankCount} steps</span>
            </button>
          </div>

          <div className="rams-logic-allocation-grid">
            <aside className="rams-logic-bank-panel">
              <div className="rams-logic-heading">
                <h3>{logicSection === "risk" ? "Risk bank" : "Method bank"}</h3>
                <button className="ghost-button" type="button" onClick={addRamsCard}>Add new</button>
              </div>
              <p className="rams-logic-help">Pick one item, then set where it appears using the trigger groups on the right.</p>
              <div className="rams-bank-card-list">
                {sectionCards.map(([cardId, card]) => {
                  const tagCount = ramsLogicDraft.optionGroups.reduce(
                    (count, group) => count + group.options.filter((option) => option.cardIds.includes(cardId)).length,
                    0
                  );
                  const residual = getRamsLcr(card, "residual");
                  const band = getRamsRiskBand(residual.rating);
                  return (
                    <button
                      key={cardId}
                      type="button"
                      className={`rams-bank-card type-${String(card.type || "risk").toLowerCase()} ${cardId === activeCardId ? "active" : ""}`}
                      onClick={() => setSelectedCardId(cardId)}
                    >
                      <strong>{card.title}</strong>
                      <span>{tagCount} trigger{tagCount === 1 ? "" : "s"}</span>
                      {card.type !== "Method" ? <small className={`rams-mini-risk ${band.className}`}>{band.code}</small> : null}
                    </button>
                  );
                })}
              </div>
            </aside>

            <section className="rams-logic-editor-panel rams-logic-allocation-editor">
              <div className="rams-logic-heading">
                <h3>Edit selected {logicSection === "risk" ? "risk" : "method"}</h3>
                <button className="ghost-button" type="button" onClick={addRamsCard}>Add new {logicSection === "risk" ? "risk" : "method"}</button>
              </div>

              {activeCard ? (
                <article className={`rams-logic-card rams-logic-editor-card type-${String(activeCard.type || "risk").toLowerCase()}`}>
                  <div className="rams-editor-topline">
                    <span className="rams-type-chip">{activeCard.type}</span>
                    <label className="rams-check">
                      <input
                        type="checkbox"
                        checked={ramsLogicDraft.baseCardIds.includes(activeCardId)}
                        onChange={() => toggleBaseRamsCard(activeCardId)}
                      />
                      Always include on every RAMS, no matter what job tags are chosen
                    </label>
                  </div>
                  <label className="rams-field-wide">
                    {activeCard.type === "Method" ? "Method title" : "Hazard"}
                    <input value={activeCard.title || ""} onChange={(event) => updateRamsCard(activeCardId, "title", event.target.value)} />
                  </label>
                  {activeCard.type === "Method" ? (
                    <>
                      <label className="rams-field-wide">
                        Trigger note
                        <input value={activeCard.trigger || ""} onChange={(event) => updateRamsCard(activeCardId, "trigger", event.target.value)} />
                      </label>
                      <label className="rams-field-wide">
                        Method statement lines
                        <textarea
                          value={(Array.isArray(activeCard.content) ? activeCard.content : []).join("\n")}
                          onChange={(event) => updateRamsCard(activeCardId, "content", event.target.value)}
                        />
                      </label>
                    </>
                  ) : (
                    <>
                      <div className="rams-logic-card-row rams-editor-two-col">
                        <label>
                          Who is at risk
                          <textarea
                            className="rams-small-textarea"
                            value={activeCard.whoAtRisk || ""}
                            onChange={(event) => updateRamsCard(activeCardId, "whoAtRisk", event.target.value)}
                          />
                        </label>
                        <label>
                          Responsibility
                          <input
                            value={activeCard.responsibility || "Installers from selected job"}
                            onChange={(event) => updateRamsCard(activeCardId, "responsibility", event.target.value)}
                          />
                        </label>
                      </div>
                      <div className="rams-risk-score-editor">
                        <div>
                          <strong>Initial LCR</strong>
                          <div className="rams-score-inputs">
                            <label>
                              L
                              <input type="number" min="0" max="5" value={activeCard.initialLikelihood ?? 2} onChange={(event) => updateRamsCard(activeCardId, "initialLikelihood", event.target.value)} />
                            </label>
                            <label>
                              C
                              <input type="number" min="0" max="5" value={activeCard.initialConsequence ?? 3} onChange={(event) => updateRamsCard(activeCardId, "initialConsequence", event.target.value)} />
                            </label>
                            <span>R {calculateRamsRisk(activeCard.initialLikelihood, activeCard.initialConsequence)}</span>
                          </div>
                        </div>
                        <div>
                          <strong>Residual LCR</strong>
                          <div className="rams-score-inputs">
                            <label>
                              L
                              <input type="number" min="0" max="5" value={activeCard.residualLikelihood ?? 1} onChange={(event) => updateRamsCard(activeCardId, "residualLikelihood", event.target.value)} />
                            </label>
                            <label>
                              C
                              <input type="number" min="0" max="5" value={activeCard.residualConsequence ?? 1} onChange={(event) => updateRamsCard(activeCardId, "residualConsequence", event.target.value)} />
                            </label>
                            <span>R {calculateRamsRisk(activeCard.residualLikelihood, activeCard.residualConsequence)}</span>
                          </div>
                        </div>
                        <div className={`rams-risk-band ${getRamsRiskBand(calculateRamsRisk(activeCard.residualLikelihood, activeCard.residualConsequence)).className}`}>
                          <strong>{getRamsRiskBand(calculateRamsRisk(activeCard.residualLikelihood, activeCard.residualConsequence)).code}</strong>
                          <span>{getRamsRiskBand(calculateRamsRisk(activeCard.residualLikelihood, activeCard.residualConsequence)).label}</span>
                        </div>
                      </div>
                      <label className="rams-field-wide">
                        Required control measure
                        <textarea
                          className="rams-control-textarea"
                          value={activeCard.controlMeasure || (Array.isArray(activeCard.content) ? activeCard.content.join("\n") : "")}
                          onChange={(event) => updateRamsCard(activeCardId, "controlMeasure", event.target.value)}
                        />
                      </label>
                    </>
                  )}
                  <div className="rams-editor-tags">
                    <div className="rams-trigger-helper">
                      <strong>When should this appear?</strong>
                      <p>Tick every job situation that should pull this {logicSection === "risk" ? "risk" : "method"} into a RAMS. Overlap is okay: if any ticked situation matches, it appears once.</p>
                    </div>
                    {linkedTagLabels.length ? (
                      <div className="rams-trigger-chip-list">
                        {linkedTagLabels.map((tag) => (
                          <button key={tag.key} type="button" onClick={() => toggleOptionCard(tag.groupIndex, tag.optionIndex, activeCardId)}>
                            {tag.label}
                          </button>
                        ))}
                      </div>
                    ) : (
                      <span>No tags linked yet. Tick tags below or make it always included.</span>
                    )}
                    {linkedTagLabels.length ? (
                      <button className="text-button danger rams-clear-tags" type="button" onClick={() => clearCardTags(activeCardId)}>
                        Clear all triggers for this item
                      </button>
                    ) : null}
                    <div className="rams-tag-matrix rams-tag-matrix-large">
                      {ramsLogicDraft.optionGroups.map((group, groupIndex) => (
                        <fieldset key={`tag-group-${group.key}`}>
                          <legend>
                            <span>{group.label}</span>
                            <span className="rams-tag-actions">
                              <button type="button" onClick={() => setGroupCardLinks(groupIndex, activeCardId, true)}>All</button>
                              <button type="button" onClick={() => setGroupCardLinks(groupIndex, activeCardId, false)}>None</button>
                            </span>
                          </legend>
                          {group.options.map((option, optionIndex) => (
                            <label key={`${group.key}-${option.value}-${optionIndex}`}>
                              <input
                                type="checkbox"
                                checked={option.cardIds.includes(activeCardId)}
                                onChange={() => toggleOptionCard(groupIndex, optionIndex, activeCardId)}
                              />
                              {option.label}
                            </label>
                          ))}
                        </fieldset>
                      ))}
                    </div>
                  </div>
                  <button className="text-button danger" type="button" onClick={() => removeRamsCard(activeCardId)}>
                    Delete {activeCard.type === "Method" ? "method" : "risk"}
                  </button>
                </article>
              ) : (
                <p className="rams-logic-empty">Select or add a card.</p>
              )}
            </section>
          </div>

          <details className="rams-tag-admin-drawer">
            <summary>Manage the job tag buttons</summary>
            <div className="rams-tag-admin-grid">
              <aside className="rams-logic-picker">
                <h3>Tag group</h3>
                {ramsLogicDraft.optionGroups.map((group, groupIndex) => (
                  <button
                    key={group.key}
                    type="button"
                    className={`rams-logic-picker-button ${groupIndex === activeGroupIndex ? "active" : ""}`}
                    onClick={() => {
                      setSelectedGroupIndex(groupIndex);
                      setSelectedOptionIndex(0);
                    }}
                  >
                    <strong>{group.label}</strong>
                    <span>{group.options.length} option{group.options.length === 1 ? "" : "s"}</span>
                  </button>
                ))}
              </aside>

              <section className="rams-logic-choice-panel">
                <div className="rams-logic-heading">
                  <h3>Buttons in this group</h3>
                  {activeGroup ? (
                    <button className="ghost-button" type="button" onClick={() => addRamsOption(activeGroupIndex)}>
                      Add button
                    </button>
                  ) : null}
                </div>
                {activeGroup ? (
                  <div className="rams-logic-choice-list">
                    {activeGroup.options.map((option, optionIndex) => (
                      <button
                        key={`${activeGroup.key}-${option.value}-${optionIndex}`}
                        type="button"
                        className={`rams-logic-choice ${optionIndex === activeOptionIndex ? "active" : ""}`}
                        onClick={() => setSelectedOptionIndex(optionIndex)}
                      >
                        <strong>{option.label}</strong>
                        <span>{option.cardIds.length} linked item{option.cardIds.length === 1 ? "" : "s"}</span>
                      </button>
                    ))}
                  </div>
                ) : null}
              </section>

              {activeGroup && activeOption ? (
                <article className="rams-logic-selected-option">
                  <h4>Edit selected tag button</h4>
                  <div className="rams-logic-option-fields">
                    <label>
                      Button text
                      <input
                        value={activeOption.label}
                        onChange={(event) => updateRamsOption(activeGroupIndex, activeOptionIndex, "label", event.target.value)}
                      />
                    </label>
                    <label>
                      Value
                      <input
                        value={activeOption.value}
                        onChange={(event) => updateRamsOption(activeGroupIndex, activeOptionIndex, "value", event.target.value)}
                      />
                    </label>
                  </div>
                  <button className="text-button danger" type="button" onClick={() => removeRamsOption(activeGroupIndex, activeOptionIndex)}>
                    Remove this button
                  </button>
                </article>
              ) : null}
            </div>
          </details>
        </section>
      </div>
    </div>
  );
}

function RamsPage({ currentUser, onLogout, notifications }) {
  const [jobs, setJobs] = useState([]);
  const [loadingJobs, setLoadingJobs] = useState(true);
  const [jobError, setJobError] = useState("");
  const [saveStatus, setSaveStatus] = useState("");
  const [savingRams, setSavingRams] = useState(false);
  const [savedRamsId, setSavedRamsId] = useState("");
  const [questions, setQuestions] = useState(RAMS_DEFAULT_QUESTIONS);
  const [ramsLogic, setRamsLogic] = useState(() => getStoredRamsLogic());
  const [ramsLogicDraft, setRamsLogicDraft] = useState(() => getStoredRamsLogic());
  const [logicStatus, setLogicStatus] = useState("");
  const [cardOrder, setCardOrder] = useState(() => getRamsCardIdsForQuestions(RAMS_DEFAULT_QUESTIONS, getStoredRamsLogic()));
  const [draggingCardId, setDraggingCardId] = useState("");
  const [ramsEdits, setRamsEdits] = useState({});
  const todayIso = getLocalTodayIso();
  const initialRamsParams = useMemo(() => {
    if (typeof window === "undefined") return { jobId: "", ramsId: "" };
    const params = new URLSearchParams(window.location.search);
    return {
      jobId: params.get("jobId") || "",
      ramsId: params.get("ramsId") || ""
    };
  }, []);

  useEffect(() => {
    let active = true;

    async function loadJobs() {
      try {
        setLoadingJobs(true);
        const response = await fetch("/api/jobs");
        if (!response.ok) throw new Error("Could not load installation jobs.");
        const payload = await response.json();
        if (!active) return;
        const jobsPayload = Array.isArray(payload) ? payload : [];
        const futureJobs = jobsPayload
          .filter((job) => String(job.date || "") >= todayIso || String(job.id || "") === String(initialRamsParams.jobId || ""))
          .sort((left, right) => String(left.date || "").localeCompare(String(right.date || "")));
        setJobs(futureJobs);
        setQuestions((current) => {
          if (current.jobId || !futureJobs.length) return current;
          const requestedJob = futureJobs.find((job) => String(job.id || "") === String(initialRamsParams.jobId || ""));
          if (requestedJob) return { ...current, jobId: String(requestedJob.id || "") };
          return { ...current, jobId: String(futureJobs[0].id || "") };
        });
      } catch (error) {
        console.error(error);
        if (active) setJobError(error.message || "Could not load installation jobs.");
      } finally {
        if (active) setLoadingJobs(false);
      }
    }

    loadJobs();
    return () => {
      active = false;
    };
  }, [initialRamsParams.jobId, todayIso]);

  const selectedJob = useMemo(
    () => jobs.find((job) => String(job.id || "") === String(questions.jobId || "")) || null,
    [jobs, questions.jobId]
  );

  const suggestedCardIds = useMemo(() => getRamsCardIdsForQuestions(questions, ramsLogic), [questions, ramsLogic]);
  const ramsReference = useMemo(() => buildRamsReference(selectedJob, questions), [selectedJob, questions]);

  useEffect(() => {
    if (!selectedJob) return;
    const savedDocuments = Array.isArray(selectedJob.ramsDocuments) ? selectedJob.ramsDocuments : [];
    const requested = savedDocuments.find((entry) => String(entry.id || "") === String(initialRamsParams.ramsId || ""));
    const savedDocument = requested || savedDocuments[0] || null;
    if (!savedDocument) {
      setSavedRamsId("");
      setRamsEdits({});
      setSaveStatus("");
      return;
    }
    setSavedRamsId(savedDocument.id || "");
    setQuestions((current) => ({
      ...current,
      ...(savedDocument.questions || {}),
      jobId: String(selectedJob.id || "")
    }));
    if (Array.isArray(savedDocument.cardOrder) && savedDocument.cardOrder.length) {
      setCardOrder(savedDocument.cardOrder);
    }
    setRamsEdits(savedDocument.edits && typeof savedDocument.edits === "object" ? savedDocument.edits : {});
    setSaveStatus(`Loaded saved RAMS ${savedDocument.reference || ""}`.trim());
  }, [initialRamsParams.ramsId, selectedJob?.id]);

  useEffect(() => {
    setCardOrder((current) => {
      const existing = current.filter((cardId) => suggestedCardIds.includes(cardId));
      const missing = suggestedCardIds.filter((cardId) => !existing.includes(cardId));
      return [...existing, ...missing];
    });
  }, [suggestedCardIds]);

  function updateQuestion(key, value) {
    setQuestions((current) => ({ ...current, [key]: value }));
  }

  function toggleTool(value) {
    setQuestions((current) => {
      const existing = Array.isArray(current.tools) ? current.tools : [];
      const nextTools = existing.includes(value)
        ? existing.filter((entry) => entry !== value)
        : [...existing, value];
      return { ...current, tools: nextTools.length ? nextTools : ["hand-tools"] };
    });
  }

  function updateMultiQuestion(key, value) {
    setQuestions((current) => {
      const existing = Array.isArray(current[key]) ? current[key] : [];
      const nextValues = existing.includes(value)
        ? existing.filter((entry) => entry !== value)
        : [...existing, value];
      return { ...current, [key]: nextValues };
    });
  }

  function updateRamsLogicDraft(updater) {
    setRamsLogicDraft((current) => normalizeRamsLogic(typeof updater === "function" ? updater(current) : updater));
    setLogicStatus("");
  }

  function saveRamsLogicDraft() {
    const nextLogic = normalizeRamsLogic(ramsLogicDraft);
    setRamsLogic(nextLogic);
    setRamsLogicDraft(nextLogic);
    saveRamsLogic(nextLogic);
    setLogicStatus("RAMS logic saved.");
  }

  function restoreDefaultRamsLogic() {
    const nextLogic = normalizeRamsLogic(RAMS_DEFAULT_LOGIC);
    setRamsLogic(nextLogic);
    setRamsLogicDraft(nextLogic);
    saveRamsLogic(nextLogic);
    setLogicStatus("Default RAMS logic restored.");
  }

  function addRamsOption(groupIndex) {
    updateRamsLogicDraft((current) => {
      const optionGroups = current.optionGroups.map((group, index) => {
        if (index !== groupIndex) return group;
        const nextIndex = group.options.length + 1;
        return {
          ...group,
          options: [
            ...group.options,
            { value: `${group.key}-${Date.now()}`, label: `New option ${nextIndex}`, cardIds: [] }
          ]
        };
      });
      return { ...current, optionGroups };
    });
  }

  function updateRamsOption(groupIndex, optionIndex, key, value) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) => {
        if (currentGroupIndex !== groupIndex) return group;
        return {
          ...group,
          options: group.options.map((option, currentOptionIndex) =>
            currentOptionIndex === optionIndex ? { ...option, [key]: value } : option
          )
        };
      })
    }));
  }

  function removeRamsOption(groupIndex, optionIndex) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) =>
        currentGroupIndex === groupIndex
          ? { ...group, options: group.options.filter((_, currentOptionIndex) => currentOptionIndex !== optionIndex) }
          : group
      )
    }));
  }

  function toggleOptionCard(groupIndex, optionIndex, cardId) {
    updateRamsLogicDraft((current) => ({
      ...current,
      optionGroups: current.optionGroups.map((group, currentGroupIndex) => {
        if (currentGroupIndex !== groupIndex) return group;
        return {
          ...group,
          options: group.options.map((option, currentOptionIndex) => {
            if (currentOptionIndex !== optionIndex) return option;
            const cardIds = option.cardIds.includes(cardId)
              ? option.cardIds.filter((entry) => entry !== cardId)
              : [...option.cardIds, cardId];
            return { ...option, cardIds };
          })
        };
      })
    }));
  }

  function addRamsCard() {
    const cardId = `custom-${Date.now()}`;
    updateRamsLogicDraft((current) => ({
      ...current,
      cards: {
        ...current.cards,
        [cardId]: {
          title: "New RAMS card",
          type: "Risk",
          trigger: "Custom",
          initialRisk: "Medium",
          residualRisk: "Low",
          content: ["Write the control measure or method step here."]
        }
      }
    }));
  }

  function updateRamsCard(cardId, key, value) {
    updateRamsLogicDraft((current) => ({
      ...current,
      cards: {
        ...current.cards,
        [cardId]: {
          ...current.cards[cardId],
          [key]: key === "content" ? String(value).split("\n").map((line) => line.trim()).filter(Boolean) : value
        }
      }
    }));
  }

  function removeRamsCard(cardId) {
    updateRamsLogicDraft((current) => {
      const nextCards = { ...current.cards };
      delete nextCards[cardId];
      return {
        ...current,
        cards: nextCards,
        baseCardIds: current.baseCardIds.filter((entry) => entry !== cardId),
        optionGroups: current.optionGroups.map((group) => ({
          ...group,
          options: group.options.map((option) => ({
            ...option,
            cardIds: option.cardIds.filter((entry) => entry !== cardId)
          }))
        }))
      };
    });
  }

  function toggleBaseRamsCard(cardId) {
    updateRamsLogicDraft((current) => ({
      ...current,
      baseCardIds: current.baseCardIds.includes(cardId)
        ? current.baseCardIds.filter((entry) => entry !== cardId)
        : [...current.baseCardIds, cardId]
    }));
  }

  function getRamsEdit(key, fallback = "") {
    return ramsEdits[key] ?? String(fallback || "");
  }

  function updateRamsEdit(key, value) {
    setRamsEdits((current) => ({ ...current, [key]: value }));
  }

  function renderEditable(key, fallback, className = "") {
    return (
      <span
        className={`rams-editable ${className}`.trim()}
        contentEditable
        suppressContentEditableWarning
        onBlur={(event) => updateRamsEdit(key, event.currentTarget.textContent || "")}
      >
        {getRamsEdit(key, fallback)}
      </span>
    );
  }

  function moveCard(cardId, direction) {
    setCardOrder((current) => {
      const index = current.indexOf(cardId);
      const nextIndex = index + direction;
      if (index < 0 || nextIndex < 0 || nextIndex >= current.length) return current;
      const next = [...current];
      const [card] = next.splice(index, 1);
      next.splice(nextIndex, 0, card);
      return next;
    });
  }

  function handleCardDrop(targetCardId) {
    if (!draggingCardId || draggingCardId === targetCardId) return;
    setCardOrder((current) => {
      const next = current.filter((cardId) => cardId !== draggingCardId);
      const targetIndex = next.indexOf(targetCardId);
      next.splice(targetIndex < 0 ? next.length : targetIndex, 0, draggingCardId);
      return next;
    });
    setDraggingCardId("");
  }

  async function saveCurrentRams() {
    if (!selectedJob) {
      setSaveStatus("Select an installation job before saving.");
      return;
    }
    try {
      setSavingRams(true);
      setSaveStatus("");
      const documentId = savedRamsId || (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : `rams-${Date.now()}`);
      const existingDocument = Array.isArray(selectedJob.ramsDocuments)
        ? selectedJob.ramsDocuments.find((entry) => String(entry.id || "") === String(documentId))
        : null;
      const body = {
        id: documentId,
        reference: getRamsEdit("reference", ramsReference),
        createdAt: existingDocument?.createdAt || new Date().toISOString(),
        questions,
        cardOrder,
        edits: ramsEdits
      };
      const response = await fetch(`/api/jobs/${encodeURIComponent(selectedJob.id)}/rams`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not save RAMS.");
      if (Array.isArray(payload.jobs)) setJobs(payload.jobs.filter((job) => String(job.date || "") >= todayIso || String(job.id || "") === String(selectedJob.id || "")));
      setSavedRamsId(payload.ramsDocument?.id || documentId);
      setSaveStatus("RAMS saved to this installation job.");
    } catch (error) {
      console.error(error);
      setSaveStatus(error.message || "Could not save RAMS.");
    } finally {
      setSavingRams(false);
    }
  }

  const selectedCards = cardOrder
    .map((cardId) => ({ id: cardId, ...ramsLogic.cards[cardId] }))
    .filter((card) => card.title);
  const methodCards = selectedCards.filter((card) => card.type === "Method");
  const riskCards = selectedCards.filter((card) => card.type !== "Method");
  const selectedActivity = ramsLogic.optionGroups.find((group) => group.key === "activity")?.options.find((option) => option.value === questions.activity)?.label || questions.activity;
  const selectedAccess = ramsLogic.optionGroups.find((group) => group.key === "access")?.options.find((option) => option.value === questions.access)?.label || questions.access;
  const selectedWorkArea = ramsLogic.optionGroups.find((group) => group.key === "workArea")?.options.find((option) => option.value === questions.workArea)?.label || questions.workArea;
  const selectedTools = (ramsLogic.optionGroups.find((group) => group.key === "tools")?.options || []).filter((option) => questions.tools.includes(option.value)).map((option) => option.label);
  const activeSavedRams = Array.isArray(selectedJob?.ramsDocuments)
    ? selectedJob.ramsDocuments.find((entry) => String(entry.id || "") === String(savedRamsId || ""))
    : null;
  const displayedInstallDate = getRamsEdit("installDate", formatJobDate(selectedJob?.date));
  const displayedCreatedDate = getRamsEdit("createdDate", formatRamsCreatedDate(activeSavedRams?.createdAt));
  const displayedOperatives = getRamsEdit("operatives", questions.operatives);
  const displayedDuration = getRamsEdit("duration", questions.duration);
  const displayedInstallers = getRamsEdit("installers", selectedJob ? getInstallerNamesForRams(selectedJob) : "To be allocated");
  const displayedActivity = getRamsEdit("activity", selectedActivity);
  const displayedAccess = getRamsEdit("access", selectedAccess);
  const displayedWorkArea = getRamsEdit("workArea", selectedWorkArea);
  const displayedTools = getRamsEdit("tools", selectedTools.join(", "));

  return (
    <div className="app-shell rams-shell">
      <div className="page rams-page">
        <MainNavBar
          currentUser={currentUser}
          active="rams"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel rams-panel">
          <div className="rams-header">
            <div>
              <span className="panel-kicker">Admin module</span>
              <h2>RAMS Builder</h2>
              <p>Select a live board job, answer the basics, then reorder the cards before printing.</p>
            </div>
            <button className="primary-button" type="button" onClick={() => window.print()} disabled={!selectedJob}>
              Print / Save PDF
            </button>
            <button className="ghost-button" type="button" onClick={saveCurrentRams} disabled={!selectedJob || savingRams}>
              {savingRams ? "Saving..." : savedRamsId ? "Save RAMS changes" : "Save RAMS to job"}
            </button>
          </div>

          {jobError ? <div className="flash error">{jobError}</div> : null}
          {saveStatus ? <div className={`flash ${saveStatus.toLowerCase().includes("could not") ? "error" : "success"}`}>{saveStatus}</div> : null}

          <div className="rams-builder-grid">
            <aside className="rams-question-panel">
              <label>
                Installation job
                <select
                  value={questions.jobId}
                  onChange={(event) => updateQuestion("jobId", event.target.value)}
                  disabled={loadingJobs || !jobs.length}
                >
                  {loadingJobs ? <option>Loading jobs...</option> : null}
                  {!loadingJobs && !jobs.length ? <option>No future jobs on the board</option> : null}
                  {jobs.map((job) => (
                    <option key={job.id} value={job.id}>
                      {formatJobDate(job.date)} - {getRamsJobTitle(job)}
                    </option>
                  ))}
                </select>
              </label>

              {ramsLogic.optionGroups.map((group) => (
                group.input === "select" && !group.multi ? (
                  <label key={group.key}>
                    {group.label}
                    <select value={questions[group.key] || ""} onChange={(event) => updateQuestion(group.key, event.target.value)}>
                      {group.options.map((option) => (
                        <option key={option.value} value={option.value}>{option.label}</option>
                      ))}
                    </select>
                  </label>
                ) : (
                  <div key={group.key} className="rams-question-group">
                    <span>{group.label}</span>
                    <div className={group.multi ? "rams-check-grid" : "rams-option-stack"}>
                      {group.options.map((option) => (
                        group.multi ? (
                          <label key={option.value} className="rams-check">
                            <input
                              type="checkbox"
                              checked={Array.isArray(questions[group.key]) && questions[group.key].includes(option.value)}
                              onChange={() => updateMultiQuestion(group.key, option.value)}
                            />
                            {option.label}
                          </label>
                        ) : (
                          <button
                            key={option.value}
                            className={`rams-choice ${questions[group.key] === option.value ? "active" : ""}`}
                            type="button"
                            onClick={() => updateQuestion(group.key, option.value)}
                          >
                            {option.label}
                          </button>
                        )
                      ))}
                    </div>
                  </div>
                )
              ))}

              <div className="rams-mini-fields">
                <label>
                  Operatives
                  <input value={questions.operatives} onChange={(event) => updateQuestion("operatives", event.target.value)} />
                </label>
                <label>
                  Duration
                  <input value={questions.duration} onChange={(event) => updateQuestion("duration", event.target.value)} />
                </label>
              </div>

              <label>
                Welfare
                <textarea value={questions.welfare} onChange={(event) => updateQuestion("welfare", event.target.value)} />
              </label>
              <label>
                Emergency arrangements
                <textarea value={questions.emergency} onChange={(event) => updateQuestion("emergency", event.target.value)} />
              </label>
              <label>
                Extra notes
                <textarea value={questions.notes} onChange={(event) => updateQuestion("notes", event.target.value)} placeholder="Anything unusual about the site or task." />
              </label>

              <button className="ghost-button rams-logic-launch" type="button" onClick={() => window.location.assign("/rams/logic")}>
                Open RAMS logic editor
              </button>
            </aside>

            <main className="rams-main">
              <section className="rams-document-preview">
                <div className="rams-doc-title">
                  <img src="/branding/signs-express-logo.svg" alt="Signs Express" />
                  <div>
                    <h3>Risk Assessment and Method Statement</h3>
                    <p>{renderEditable("jobTitle", selectedJob ? getRamsJobTitle(selectedJob) : "Select a job to generate the document")} - Ref: {renderEditable("reference", ramsReference)}</p>
                  </div>
                </div>
                <div className="rams-doc-meta">
                  <span><strong>Installation date:</strong> {renderEditable("installDate", displayedInstallDate)}</span>
                  <span><strong>RAMS created:</strong> {renderEditable("createdDate", displayedCreatedDate)}</span>
                  <span><strong>Operatives:</strong> {renderEditable("operatives", displayedOperatives)}</span>
                  <span><strong>Duration:</strong> {renderEditable("duration", displayedDuration)}</span>
                  <span><strong>Installers:</strong> {renderEditable("installers", displayedInstallers)}</span>
                  <span><strong>Activity:</strong> {renderEditable("activity", displayedActivity)}</span>
                  <span><strong>Access:</strong> {renderEditable("access", displayedAccess)}</span>
                  <span><strong>Work area:</strong> {renderEditable("workArea", displayedWorkArea)}</span>
                  <span><strong>Tools:</strong> {renderEditable("tools", displayedTools)}</span>
                </div>
                <div className="rams-doc-section">
                  <h4>Site Details</h4>
                  <p><strong>Customer:</strong> {renderEditable("customerName", selectedJob?.customerName || "-")}</p>
                  <p><strong>Site address:</strong> {renderEditable("siteAddress", selectedJob ? getRamsJobAddress(selectedJob) : "-")}</p>
                  <p><strong>Contact name:</strong> {renderEditable("contactName", selectedJob?.contact || "-")}</p>
                  <p><strong>Contact number:</strong> {renderEditable("contactNumber", selectedJob?.number || "-")}</p>
                  <p><strong>Scope:</strong> {renderEditable("scope", selectedJob?.description || selectedActivity, "wide-edit")}</p>
                </div>
                <div className="rams-doc-section">
                  <h4>Arrangements</h4>
                  <p><strong>Welfare:</strong> {renderEditable("welfare", questions.welfare)}</p>
                  <p><strong>Emergency:</strong> {renderEditable("emergency", questions.emergency)}</p>
                  <p><strong>Notes:</strong> {renderEditable("notes", questions.notes || "-")}</p>
                </div>
                <div className="rams-doc-section rams-doc-risk-section">
                  <h4>Risk Assessment</h4>
                  <div className="rams-risk-key">
                    <span><strong>L</strong> Likelihood</span>
                    <span><strong>C</strong> Consequence</span>
                    <span><strong>R</strong> Risk = L x C</span>
                    <span className="risk-none">0-4 N</span>
                    <span className="risk-low">5-10 L</span>
                    <span className="risk-medium">11-15 M</span>
                    <span className="risk-high">16+ H</span>
                  </div>
                  <div className="rams-risk-table-wrap">
                    <table className="rams-risk-table">
                      <colgroup>
                        <col className="rams-col-hazard" />
                        <col className="rams-col-harmed" />
                        <col className="rams-col-score" />
                        <col className="rams-col-score" />
                        <col className="rams-col-score" />
                        <col className="rams-col-controls" />
                        <col className="rams-col-responsibility" />
                        <col className="rams-col-score" />
                        <col className="rams-col-score" />
                        <col className="rams-col-score" />
                        <col className="rams-col-final-risk" />
                      </colgroup>
                      <thead>
                        <tr>
                          <th rowSpan="2">Hazard</th>
                          <th rowSpan="2">Who may be harmed</th>
                          <th colSpan="3">Initial risk</th>
                          <th rowSpan="2">Req'd control measure</th>
                          <th rowSpan="2">Responsibility</th>
                          <th colSpan="3">Residual risk</th>
                          <th rowSpan="2">Risk</th>
                        </tr>
                        <tr>
                          <th>L</th>
                          <th>C</th>
                          <th>R</th>
                          <th>L</th>
                          <th>C</th>
                          <th>R</th>
                        </tr>
                      </thead>
                      <tbody>
                        {riskCards.map((card) => {
                          const initial = getRamsLcr(card, "initial");
                          const residual = getRamsLcr(card, "residual");
                          const riskBand = getRamsRiskBand(residual.rating);
                          const controlLines = String(card.controlMeasure || (Array.isArray(card.content) ? card.content.join("\n") : "")).split("\n").filter(Boolean);
                          const responsibility = !card.responsibility || card.responsibility === "Matt Carroll" || card.responsibility === "Installers from selected job"
                            ? displayedInstallers
                            : card.responsibility;
                          return (
                            <tr key={`risk-row-${card.id}`}>
                              <td>
                                <strong>{renderEditable(`risk-${card.id}-title`, card.title)}</strong>
                                <span>{card.type}</span>
                              </td>
                              <td>{renderEditable(`risk-${card.id}-harmed`, card.whoAtRisk || "Employees\nThird parties")}</td>
                              <td>{renderEditable(`risk-${card.id}-initial-l`, initial.likelihood)}</td>
                              <td>{renderEditable(`risk-${card.id}-initial-c`, initial.consequence)}</td>
                              <td><strong>{initial.rating}</strong></td>
                              <td>
                                <ul>
                                  {controlLines.map((line, lineIndex) => (
                                    <li key={`risk-control-${card.id}-${line}`}>{renderEditable(`risk-${card.id}-control-${lineIndex}`, line)}</li>
                                  ))}
                                </ul>
                              </td>
                              <td>{renderEditable(`risk-${card.id}-responsibility`, responsibility)}</td>
                              <td>{renderEditable(`risk-${card.id}-residual-l`, residual.likelihood)}</td>
                              <td>{renderEditable(`risk-${card.id}-residual-c`, residual.consequence)}</td>
                              <td><strong>{residual.rating}</strong></td>
                              <td className={`rams-risk-final ${riskBand.className}`}>
                                <strong>{riskBand.code}</strong>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
                <div className="rams-doc-section rams-doc-method-section">
                  <h4>Method Statement</h4>
                  {methodCards.map((card) => {
                    const cardIndex = selectedCards.findIndex((entry) => entry.id === card.id);
                    return (
                      <div
                        key={`doc-${card.id}`}
                        className="rams-doc-card"
                        draggable
                        onDragStart={() => setDraggingCardId(card.id)}
                        onDragEnd={() => setDraggingCardId("")}
                        onDragOver={(event) => event.preventDefault()}
                        onDrop={() => handleCardDrop(card.id)}
                      >
                        <div>
                          <strong>{renderEditable(`method-${card.id}-title`, card.title)}</strong>
                          <span className="rams-card-actions no-print">
                            <button type="button" className="icon-button" onClick={() => moveCard(card.id, -1)} disabled={cardIndex <= 0} aria-label="Move method up">^</button>
                            <button type="button" className="icon-button" onClick={() => moveCard(card.id, 1)} disabled={cardIndex === selectedCards.length - 1} aria-label="Move method down">v</button>
                          </span>
                        </div>
                        <ul>
                          {card.content.map((line, lineIndex) => (
                            <li key={`doc-${card.id}-${line}`}>{renderEditable(`method-${card.id}-line-${lineIndex}`, line)}</li>
                          ))}
                        </ul>
                      </div>
                    );
                  })}
                </div>
                <div className="rams-signoff">
                  <span>Prepared by: {currentUser?.displayName || "Signs Express"}</span>
                  <span>Accepted by site:</span>
                  <span>Signature:</span>
                  <span>Date:</span>
                </div>
              </section>
            </main>
          </div>
        </section>
      </div>
    </div>
  );
}

function HolidaysPage({
  currentUser,
  onLogout,
  notifications,
  holidays,
  holidayRequests,
  approvedHolidayRequests,
  holidayStaff,
  holidayAllowances,
  holidayEvents,
  holidayRows,
  holidayYearStart,
  holidayYearLabel,
  currentHolidayYearStart,
  setHolidayYearStart,
  holidayRequestOpen,
  setHolidayRequestOpen,
  holidayRequestForm,
  setHolidayRequestForm,
  holidayRequestSaving,
  holidayCancelOpen,
  setHolidayCancelOpen,
  holidayCancelForm,
  setHolidayCancelForm,
  holidayEventOpen,
  setHolidayEventOpen,
  holidayEventForm,
  setHolidayEventForm,
  holidayEventSaving,
  holidayAllowanceSavingKey,
  onChangeHolidayAllowanceDraft,
  onSaveHolidayAllowance,
  onToggleHolidayDate,
  onSubmitHolidayEvent,
  onDeleteHolidayEvent,
  onSubmitHolidayRequest,
  onReviewHolidayRequest,
  onCancelHolidayRequest
}) {
  const canReview = canEditHolidays(currentUser);
  const currentPerson = getHolidayStaffPersonForUser(currentUser);
  const showingFutureYear = holidayYearStart > currentHolidayYearStart;
  const [selectedHolidayPerson, setSelectedHolidayPerson] = useState("");
  const cancellableHolidayRequests = useMemo(() => {
    if (canReview || !currentPerson) return [];
    const personKey = getHolidayStaffIdentityKey(currentPerson);
    const currentUserId = String(currentUser?.id || "");
    return approvedHolidayRequests.filter((request) => {
      const requestStatus = String(request.status || "").trim().toLowerCase();
      const requestAction = String(request.action || "book").trim().toLowerCase();
      const sameUser =
        (currentUserId && String(request.requestedByUserId || "") === currentUserId) ||
        getHolidayStaffIdentityKey(request.person) === personKey;
      return sameUser && requestStatus === "approved" && requestAction === "book";
    });
  }, [canReview, currentPerson, currentUser, approvedHolidayRequests]);
  const visibleHolidayAllowances = useMemo(
    () => {
      const allowedPeople = new Set(
        holidayStaff.map((entry) => getHolidayStaffIdentityKey(entry.person || entry.fullName || entry.name))
      );
      return holidayAllowances
        .map((rawEntry) => getHolidayAllowanceSummary(rawEntry))
        .filter((entry) => allowedPeople.has(getHolidayStaffIdentityKey(entry.person)));
    },
    [holidayAllowances, holidayStaff]
  );
  const activeHolidayFilter = selectedHolidayPerson;
  const filteredHolidayRows = useMemo(
    () =>
      holidayRows.map((month) => ({
        ...month,
        days: month.days.map((day) => ({
          ...day,
          holidays: activeHolidayFilter
            ? day.holidays.filter(
                (holiday) =>
                  getHolidayStaffIdentityKey(holiday.person) === getHolidayStaffIdentityKey(activeHolidayFilter)
              )
            : day.holidays
        }))
      })),
    [holidayRows, activeHolidayFilter]
  );

  return (
    <div className="app-shell holidays-shell">
      <div className="page holidays-page">
        <MainNavBar
          currentUser={currentUser}
          active="holidays"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel holidays-panel">
          <div className="holidays-toolbar">
            <div>
              <h2>Holiday Calendar {holidayYearLabel}</h2>
            </div>
            <div className="holidays-toolbar-actions">
              {showingFutureYear ? (
                <button className="ghost-button" type="button" onClick={() => setHolidayYearStart(currentHolidayYearStart)}>
                  Current year
                </button>
                ) : null}
                <button className="ghost-button" type="button" onClick={() => setHolidayRequestOpen(true)}>
                  Request holiday
                </button>
                {!canReview ? (
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={() => setHolidayCancelOpen(true)}
                    disabled={!cancellableHolidayRequests.length}
                  >
                    Cancel holiday
                  </button>
                ) : null}
              </div>
            </div>

          {holidayRequests.length ? (
            <section className="holiday-requests-panel">
              <div className="holiday-requests-head">
                <h3>{canReview ? "Pending Requests" : "Your Requests"}</h3>
              </div>
              <div className="holiday-request-list">
                {holidayRequests.map((request) => (
                  <article key={request.id} className={`holiday-request-card status-${request.status || "pending"}`}>
                    <div className="holiday-request-main">
                      <strong>{request.person}</strong>
                      <span>{formatHolidayRequestDateRange(request.startDate, request.endDate)}</span>
                      <span>{request.duration || "Full Day"}</span>
                      <span>{String(request.action || "book").toLowerCase() === "cancel" ? "Cancellation request" : "Holiday request"}</span>
                      {request.notes ? <p>{request.notes}</p> : null}
                    </div>
                    <div className="holiday-request-side">
                      <span className={`holiday-request-status status-${request.status || "pending"}`}>
                        {request.status || "pending"}
                      </span>
                      <small>Requested by {request.requestedByName || request.person}</small>
                      {canReview && request.status === "pending" ? (
                        <div className="holiday-request-actions">
                          <button className="ghost-button" type="button" onClick={() => onReviewHolidayRequest(request.id, "approved")}>
                            Approve
                          </button>
                          <button className="text-button danger" type="button" onClick={() => onReviewHolidayRequest(request.id, "rejected")}>
                            Reject
                          </button>
                        </div>
                      ) : null}
                    </div>
                  </article>
                ))}
              </div>
            </section>
          ) : null}

          <div className="holiday-calendar-wrap">
            <div className="holiday-calendar-tools">
              <div className="holiday-calendar-filter-summary">
                {activeHolidayFilter ? (
                  <>
                    <span className="holiday-filter-label">Showing</span>
                    <button
                      type="button"
                      className="holiday-filter-pill active"
                      onClick={() => setSelectedHolidayPerson("")}
                    >
                      {activeHolidayFilter}
                    </button>
                    <button
                      type="button"
                      className="text-button"
                      onClick={() => setSelectedHolidayPerson("")}
                    >
                      Show everyone
                    </button>
                  </>
                ) : (
                  <span className="holiday-filter-label">
                    {canReview ? "Showing all employees" : "Showing all employees on the calendar"}
                  </span>
                )}
              </div>
            </div>
            <div className="holiday-calendar-grid">
              <div className="holiday-calendar-header month-label-cell">Month</div>
              {Array.from({ length: 31 }, (_, index) => (
                <div key={`header-${index + 1}`} className="holiday-calendar-header day-header-cell">
                  {String(index + 1).padStart(2, "0")}
                </div>
              ))}

              {filteredHolidayRows.map((month) => (
                <div key={month.id} className="holiday-calendar-row">
                  <div className="month-label-cell month-row-label">{month.label}</div>
                  {month.days.map((day) => (
                      (() => {
                        const matchingHoliday = activeHolidayFilter
                          ? day.holidays.find(
                              (holiday) =>
                                getHolidayStaffIdentityKey(holiday.person) ===
                                getHolidayStaffIdentityKey(activeHolidayFilter)
                            )
                          : null;

                      return (
                        <div
                          key={day.key}
                          className={[
                            "holiday-day-cell",
                            !day.inMonth ? "is-empty" : "",
                            day.weekend ? "is-weekend" : "",
                            day.bankHoliday ? "is-bank-holiday" : "",
                            canReview && activeHolidayFilter && day.inMonth && !day.weekend && !day.bankHoliday ? "is-editable" : ""
                          ].join(" ").trim()}
                          title={day.bankHoliday || day.isoDate || ""}
                            onClick={() => {
                              if (!canReview || !day.inMonth) return;
                              if (activeHolidayFilter) {
                                if (day.weekend || day.bankHoliday) return;
                                if (matchingHoliday && isBirthdayHoliday(matchingHoliday)) return;
                                onToggleHolidayDate(day.isoDate, activeHolidayFilter);
                                return;
                              }
                              setHolidayEventForm({
                                id: day.events?.[0]?.id || "",
                                date: day.isoDate,
                                title: day.events?.[0]?.title || ""
                              });
                              setHolidayEventOpen(true);
                            }}
                          >
                            {day.events.map((event) => (
                              <button
                                key={`${day.key}-event-${event.id}`}
                                type="button"
                                className="holiday-day-event"
                                onClick={(clickEvent) => {
                                  clickEvent.stopPropagation();
                                  if (!canReview) return;
                                  setHolidayEventForm({
                                    id: event.id,
                                    date: day.isoDate,
                                    title: event.title || ""
                                  });
                                  setHolidayEventOpen(true);
                                }}
                              >
                                {event.title}
                              </button>
                            ))}
                            {day.holidays.map((holiday) => (
                              <span
                                key={`${day.key}-${holiday.id}`}
                              className={`holiday-day-token ${HOLIDAY_PERSON_COLORS[holiday.person] || "holiday-person-black"} ${isBirthdayHoliday(holiday) ? "holiday-birthday-token" : ""}`}
                            >
                              {getHolidayDisplayToken(holiday.person)}
                              {holiday.duration === "Morning" ? " AM" : holiday.duration === "Afternoon" ? " PM" : ""}
                            </span>
                          ))}
                        </div>
                      );
                    })()
                  ))}
                </div>
              ))}
            </div>
          </div>

          <section className="holiday-breakdown-panel">
            <div className="holiday-requests-head">
                <h3>{canReview ? `Holiday Breakdown ${holidayYearLabel}` : `Your Holiday Breakdown ${holidayYearLabel}`}</h3>
              </div>
            <div className="holiday-breakdown-wrap">
              <table className="holiday-breakdown-table">
                <colgroup>
                  <col className="holiday-col-employee" />
                  <col span="11" className="holiday-col-metric" />
                </colgroup>
                <thead>
                  <tr>
                    <th>Employee</th>
                    <th>Birthday</th>
                    <th>Number Days work pw</th>
                    <th>Standard Entitlement (21 + 8BH)</th>
                    <th>Extra Days (Service)</th>
                    <th>Pro-rata Allowance</th>
                    <th>Allocated for Xmas</th>
                    <th>Allocated Bank Holiday</th>
                    <th>Total Days Booked</th>
                    <th>Days Left to Book</th>
                    <th>Unpaid Days Booked</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleHolidayAllowances.map((entry) => {
                    const isSelected =
                      String(activeHolidayFilter || "").trim().toLowerCase() ===
                      String(entry.person || "").trim().toLowerCase();
                    return (
                    <tr key={`${holidayYearStart}-${entry.person}`}>
                      <td>
                        <button
                          type="button"
                          className={`holiday-person-filter ${isSelected ? "active" : ""}`}
                          onClick={() =>
                            setSelectedHolidayPerson((current) =>
                              current.toLowerCase() === String(entry.person || "").trim().toLowerCase() ? "" : entry.person
                            )
                          }
                        >
                          {entry.fullName}
                        </button>
                      </td>
                      <td>
                        {canReview ? (
                          <input
                            className="holiday-allowance-input holiday-birthday-input"
                            type="date"
                            value={entry.birthDate || ""}
                            disabled={holidayAllowanceSavingKey === `${entry.person}:birthDate`}
                            onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { birthDate: event.target.value })}
                            onBlur={(event) => onSaveHolidayAllowance(entry.person, { birthDate: event.target.value })}
                          />
                        ) : (
                          <span className="holiday-birthday-label">{formatHolidayBirthday(entry.birthDate)}</span>
                        )}
                      </td>
                      {[
                        ["workDaysPerWeek", entry.workDaysPerWeek],
                        ["standardEntitlement", entry.standardEntitlement],
                        ["extraServiceDays", entry.extraServiceDays]
                      ].map(([field, value]) => (
                        <td key={`${entry.person}-${field}`}>
                          {canReview ? (
                            <input
                              className="holiday-allowance-input"
                              type="number"
                              min="0"
                              step="0.5"
                              value={value}
                              disabled={holidayAllowanceSavingKey === `${entry.person}:${field}`}
                              onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { [field]: event.target.value })}
                              onBlur={(event) => onSaveHolidayAllowance(entry.person, { [field]: event.target.value })}
                              onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                  event.preventDefault();
                                  event.currentTarget.blur();
                                }
                              }}
                            />
                          ) : (
                            <span>{value}</span>
                          )}
                        </td>
                      ))}
                      <td><strong>{entry.prorataAllowance}</strong></td>
                      {[
                        ["christmasDays", entry.christmasDays],
                        ["bankHolidayDays", entry.bankHolidayDays]
                      ].map(([field, value]) => (
                        <td key={`${entry.person}-${field}`}>
                          {canReview ? (
                            <input
                              className="holiday-allowance-input"
                              type="number"
                              min="0"
                              step="0.5"
                              value={value}
                              disabled={holidayAllowanceSavingKey === `${entry.person}:${field}`}
                              onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { [field]: event.target.value })}
                              onBlur={(event) => onSaveHolidayAllowance(entry.person, { [field]: event.target.value })}
                              onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                  event.preventDefault();
                                  event.currentTarget.blur();
                                }
                              }}
                            />
                          ) : (
                            <span>{value}</span>
                          )}
                        </td>
                      ))}
                      <td>{entry.bookedDays}</td>
                      <td className={entry.daysLeft < 0 ? "holiday-days-negative" : "holiday-days-positive"}>
                        <strong>{entry.daysLeft}</strong>
                      </td>
                      <td>
                        {canReview ? (
                          <input
                            className="holiday-allowance-input"
                            type="number"
                            min="0"
                            step="0.5"
                            value={entry.unpaidDaysBooked || 0}
                            disabled={holidayAllowanceSavingKey === `${entry.person}:unpaidDaysBooked`}
                            onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { unpaidDaysBooked: event.target.value })}
                            onBlur={(event) => onSaveHolidayAllowance(entry.person, { unpaidDaysBooked: event.target.value })}
                            onKeyDown={(event) => {
                              if (event.key === "Enter") {
                                event.preventDefault();
                                event.currentTarget.blur();
                              }
                            }}
                          />
                        ) : (
                          <span>{entry.unpaidDaysBooked || 0}</span>
                        )}
                      </td>
                    </tr>
                  )})}
                </tbody>
              </table>
            </div>
          </section>

          <div className="holiday-year-footer">
            <button className="ghost-button" type="button" onClick={() => setHolidayYearStart((current) => current + 1)}>
              {holidayYearStart + 1} Holidays
            </button>
          </div>
        </section>
      </div>

        {holidayRequestOpen ? (
          <div className="modal-backdrop" onClick={() => setHolidayRequestOpen(false)}>
          <div className="modal holiday-request-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>Request holiday</h3>
                <p>Send a request for approval. Approved holidays will show on the board automatically.</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setHolidayRequestOpen(false)}>
                x
              </button>
            </div>

            <form
              className="job-form"
              onSubmit={(event) => {
                event.preventDefault();
                onSubmitHolidayRequest();
              }}
            >
              {canReview ? (
                <label>
                  Employee
                  <select
                    value={holidayRequestForm.person}
                    onChange={(event) => setHolidayRequestForm((current) => ({ ...current, person: event.target.value }))}
                  >
                    <option value="">Select employee</option>
                    {holidayStaff.map((entry) => (
                      <option key={entry.person} value={entry.person}>
                        {entry.code} - {entry.fullName}
                      </option>
                    ))}
                  </select>
                </label>
              ) : (
                <label>
                  Employee
                  <input type="text" value={currentPerson} readOnly />
                </label>
              )}

              <div className="split-fields">
                <label>
                  Start date
                  <input
                    type="date"
                    value={holidayRequestForm.startDate}
                    onChange={(event) =>
                      setHolidayRequestForm((current) => ({
                        ...current,
                        startDate: event.target.value,
                        endDate: current.endDate || event.target.value
                      }))
                    }
                  />
                </label>

                <label>
                  End date
                  <input
                    type="date"
                    value={holidayRequestForm.endDate}
                    onChange={(event) => setHolidayRequestForm((current) => ({ ...current, endDate: event.target.value }))}
                  />
                </label>
              </div>

              <label>
                Duration
                <select
                  value={holidayRequestForm.duration}
                  onChange={(event) => setHolidayRequestForm((current) => ({ ...current, duration: event.target.value }))}
                >
                  <option value="Full Day">Full Day</option>
                  <option value="Morning">Morning</option>
                  <option value="Afternoon">Afternoon</option>
                </select>
              </label>

              <label>
                Notes
                <textarea
                  rows="4"
                  value={holidayRequestForm.notes}
                  onChange={(event) => setHolidayRequestForm((current) => ({ ...current, notes: event.target.value }))}
                />
              </label>

              <div className="form-actions">
                <button className="primary-button" type="submit" disabled={holidayRequestSaving}>
                  {holidayRequestSaving ? "Sending..." : "Send request"}
                </button>
                <button className="ghost-button" type="button" onClick={() => setHolidayRequestOpen(false)}>
                  Cancel
                </button>
              </div>
            </form>
          </div>
          </div>
          ) : null}

        {holidayCancelOpen ? (
          <div className="modal-backdrop" onClick={() => setHolidayCancelOpen(false)}>
            <div className="modal holiday-request-modal" onClick={(event) => event.stopPropagation()}>
              <div className="modal-head">
                <div>
                  <h3>Cancel holiday</h3>
                  <p>Send a cancellation request for approval. Nothing will be removed until an admin approves it.</p>
                </div>
                <button className="icon-button" type="button" onClick={() => setHolidayCancelOpen(false)}>
                  x
                </button>
              </div>

              <form
                className="job-form"
                onSubmit={(event) => {
                  event.preventDefault();
                  onCancelHolidayRequest();
                }}
              >
                <label>
                  Approved holiday
                  <select
                    value={holidayCancelForm.requestId}
                    onChange={(event) => setHolidayCancelForm((current) => ({ ...current, requestId: event.target.value }))}
                  >
                    <option value="">Select approved holiday request</option>
                    {cancellableHolidayRequests.map((holidayRequest) => (
                      <option key={holidayRequest.id} value={holidayRequest.id}>
                        {formatHolidayRequestDateRange(holidayRequest.startDate, holidayRequest.endDate)} - {holidayRequest.duration || "Full Day"}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  Notes
                  <textarea
                    rows="4"
                    value={holidayCancelForm.notes}
                    onChange={(event) => setHolidayCancelForm((current) => ({ ...current, notes: event.target.value }))}
                  />
                </label>

                <div className="form-actions">
                  <button className="primary-button" type="submit" disabled={holidayRequestSaving || !cancellableHolidayRequests.length}>
                    {holidayRequestSaving ? "Sending..." : "Send cancellation request"}
                  </button>
                  <button className="ghost-button" type="button" onClick={() => setHolidayCancelOpen(false)}>
                    Close
                  </button>
                </div>
              </form>
            </div>
          </div>
        ) : null}

          {holidayEventOpen ? (
          <div className="modal-backdrop" onClick={() => setHolidayEventOpen(false)}>
            <div className="modal holiday-request-modal" onClick={(event) => event.stopPropagation()}>
              <div className="modal-head">
                <div>
                  <h3>Calendar event</h3>
                  <p>Add something like Christmas shutdown or Summer party. This will not affect holiday allowances.</p>
                </div>
                <button className="icon-button" type="button" onClick={() => setHolidayEventOpen(false)}>
                  x
                </button>
              </div>

              <form
                className="job-form"
                onSubmit={(event) => {
                  event.preventDefault();
                  onSubmitHolidayEvent();
                }}
              >
                <label>
                  Date
                  <input
                    type="date"
                    value={holidayEventForm.date}
                    onChange={(event) => setHolidayEventForm((current) => ({ ...current, date: event.target.value }))}
                  />
                </label>

                <label>
                  Event title
                  <input
                    type="text"
                    value={holidayEventForm.title}
                    placeholder="Christmas shutdown"
                    onChange={(event) => setHolidayEventForm((current) => ({ ...current, title: event.target.value }))}
                  />
                </label>

                <div className="modal-actions">
                  <button className="primary-button" type="submit" disabled={holidayEventSaving}>
                    {holidayEventSaving ? "Saving..." : holidayEventForm.id ? "Update event" : "Add event"}
                  </button>
                  {holidayEventForm.id ? (
                    <button
                      className="text-button danger"
                      type="button"
                      onClick={() => onDeleteHolidayEvent(holidayEventForm.id)}
                    >
                      Delete
                    </button>
                  ) : null}
                  <button className="ghost-button" type="button" onClick={() => setHolidayEventOpen(false)}>
                    Cancel
                  </button>
                </div>
              </form>
            </div>
          </div>
        ) : null}
      </div>
    );
  }

function createMileageLine(overrides = {}) {
  return {
    id: overrides.id || (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : String(Date.now())),
    date: overrides.date || "",
    from: overrides.from || "",
    to: overrides.to || "",
    note: overrides.note || "",
    miles: overrides.miles ?? ""
  };
}

function MileageUserPage({ currentUser, onLogout, notifications, onRefreshNotifications }) {
  const initialMonth = useMemo(() => {
    const params = new URLSearchParams(typeof window !== "undefined" ? window.location.search : "");
    return params.get("month") || toMonthIdFromIso(getLocalTodayIso());
  }, []);
  const [monthId, setMonthId] = useState(initialMonth);
  const [lines, setLines] = useState([createMileageLine()]);
  const [history, setHistory] = useState([]);
  const [monthLabel, setMonthLabel] = useState("");
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [estimatingLineId, setEstimatingLineId] = useState("");
  const [statusMessage, setStatusMessage] = useState(null);

  const totalMiles = useMemo(
    () => Math.round(lines.reduce((sum, line) => sum + (Number(line.miles) || 0), 0) * 10) / 10,
    [lines]
  );

  function updateLine(lineId, key, value) {
    setLines((current) =>
      current.map((line) => (line.id === lineId ? { ...line, [key]: value } : line))
    );
  }

  function removeLine(lineId) {
    setLines((current) => {
      const next = current.filter((line) => line.id !== lineId);
      return next.length ? next : [createMileageLine()];
    });
  }

  function applyMileagePayload(payload) {
    setMonthLabel(payload.monthLabel || "");
    setHistory(Array.isArray(payload.history) ? payload.history : []);
    setLines([createMileageLine()]);
  }

  async function loadMileage(nextMonthId = monthId) {
    try {
      setLoading(true);
      const response = await fetch(`/api/mileage?month=${encodeURIComponent(nextMonthId)}`);
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not load mileage.");
      applyMileagePayload(payload);
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not load mileage.", "error"));
    } finally {
      setLoading(false);
    }
  }

  async function estimateLine(line) {
    if (!line.from.trim() || !line.to.trim()) {
      setStatusMessage(createMessage("Enter both From and To before estimating miles.", "error"));
      return;
    }

    try {
      setEstimatingLineId(line.id);
      const response = await fetch("/api/mileage/estimate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ from: line.from, to: line.to })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not estimate mileage.");
      if (payload.resolved && Number(payload.miles) > 0) {
        updateLine(line.id, "miles", String(payload.miles));
        setStatusMessage(createMessage("Mileage estimate added. You can still adjust it if needed.", "success"));
      } else {
        setStatusMessage(createMessage(payload.message || "Could not estimate that route. Please enter the miles manually.", "error"));
      }
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not estimate mileage.", "error"));
    } finally {
      setEstimatingLineId("");
    }
  }

  async function submitMileage(event) {
    event.preventDefault();
    const cleanLines = lines
      .map((line) => ({
        id: line.id,
        date: line.date,
        from: line.from.trim(),
        to: line.to.trim(),
        note: line.note.trim(),
        miles: Number(line.miles) || 0
      }))
      .filter((line) => line.date || line.from || line.to || line.note || line.miles);

    if (!cleanLines.length) {
      setStatusMessage(createMessage("Add at least one mileage line before submitting.", "error"));
      return;
    }
    if (cleanLines.some((line) => !line.date || !line.from || !line.to || !line.note || !line.miles)) {
      setStatusMessage(createMessage("Every journey needs Date, From, To, Miles and a note explaining what it was for.", "error"));
      return;
    }

    try {
      setSaving(true);
      const response = await fetch("/api/mileage", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ monthId, lines: cleanLines })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not submit mileage.");
      applyMileagePayload(payload);
      await onRefreshNotifications?.();
      setStatusMessage(createMessage("Mileage submitted to Matt.", "success"));
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not submit mileage.", "error"));
    } finally {
      setSaving(false);
    }
  }

  async function deleteMileageJourney(targetMonthId, lineId, targetMonthLabel = "") {
    if (!targetMonthId || !lineId) return;
    if (!window.confirm(`Delete this journey from ${targetMonthLabel || targetMonthId}?`)) return;

    try {
      setDeleting(true);
      const response = await fetch(
        `/api/mileage/${encodeURIComponent(targetMonthId)}/lines/${encodeURIComponent(lineId)}`,
        { method: "DELETE" }
      );
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not delete mileage journey.");
      if (targetMonthId === monthId) {
        applyMileagePayload(payload);
      } else {
        setHistory(Array.isArray(payload.history) ? payload.history : []);
      }
      await onRefreshNotifications?.();
      setStatusMessage(createMessage("Mileage journey deleted.", "success"));
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not delete mileage journey.", "error"));
    } finally {
      setDeleting(false);
    }
  }

  useEffect(() => {
    loadMileage(monthId);
  }, [monthId]);

  useEffect(() => {
    if (!statusMessage) return undefined;
    const timer = window.setTimeout(() => setStatusMessage(null), 3500);
    return () => window.clearTimeout(timer);
  }, [statusMessage]);

  return (
    <div className="app-shell mileage-shell">
      <div className="page mileage-page">
        <MainNavBar currentUser={currentUser} active="mileage" onLogout={onLogout} notifications={notifications} />

        <section className="panel mileage-panel">
          <div className="mileage-head">
            <div>
              <p className="eyebrow">Mileage</p>
              <h2>{monthLabel || "Mileage claim"}</h2>
              <p>Add journeys, estimate the driving miles, and submit the monthly total to Matt.</p>
            </div>
            <div className="mileage-month-tools">
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, -1))}>
                Previous month
              </button>
              <input
                type="month"
                value={monthId}
                onChange={(event) => setMonthId(event.target.value || toMonthIdFromIso(getLocalTodayIso()))}
              />
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, 1))}>
                Next month
              </button>
            </div>
          </div>

          {statusMessage ? <div className={`flash ${statusMessage.tone}`}>{statusMessage.text}</div> : null}

          <form className="mileage-form" onSubmit={submitMileage}>
            <div className="mileage-lines">
              {lines.map((line, index) => (
                <div key={line.id} className="mileage-line">
                  <span className="mileage-line-number">{index + 1}</span>
                  <label className="mileage-date-field">
                    Date
                    <input
                      type="date"
                      required
                      value={line.date}
                      onChange={(event) => updateLine(line.id, "date", event.target.value)}
                    />
                  </label>
                  <label>
                    From
                    <input
                      type="text"
                      required
                      value={line.from}
                      placeholder="Start destination"
                      onChange={(event) => updateLine(line.id, "from", event.target.value)}
                    />
                  </label>
                  <label>
                    To
                    <input
                      type="text"
                      required
                      value={line.to}
                      placeholder="End destination"
                      onChange={(event) => updateLine(line.id, "to", event.target.value)}
                    />
                  </label>
                  <label className="mileage-note-field">
                    Journey note
                    <input
                      type="text"
                      required
                      value={line.note}
                      placeholder="What was this journey for?"
                      onChange={(event) => updateLine(line.id, "note", event.target.value)}
                    />
                  </label>
                  <label className="mileage-miles-field">
                    Miles
                    <input
                      type="number"
                      required
                      min="0"
                      step="0.1"
                      value={line.miles}
                      placeholder="0"
                      onChange={(event) => updateLine(line.id, "miles", event.target.value)}
                    />
                  </label>
                  <div className="mileage-line-actions">
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => estimateLine(line)}
                      disabled={estimatingLineId === line.id}
                    >
                      {estimatingLineId === line.id ? "Checking..." : "Suggest"}
                    </button>
                    <button className="text-button danger" type="button" onClick={() => removeLine(line.id)}>
                      Remove
                    </button>
                  </div>
                </div>
              ))}
            </div>

            <div className="mileage-footer">
              <button
                className="ghost-button mileage-add-line"
                type="button"
                onClick={() => setLines((current) => [...current, createMileageLine()])}
              >
                + Add journey
              </button>
              <div className="mileage-total">
                <span>Total miles to submit</span>
                <strong>{totalMiles.toFixed(1)} miles</strong>
              </div>
              <button className="primary-button" type="submit" disabled={saving || loading}>
                {saving ? "Submitting..." : "Submit mileage"}
              </button>
            </div>
          </form>
        </section>

        <section className="panel mileage-history-panel">
          <div className="mileage-history-head">
            <h3>History</h3>
            <p>Your previously submitted mileage totals.</p>
          </div>
          {history.length ? (
            <div className="mileage-history-list">
              {history.map((entry) => (
                <article
                  key={entry.id || entry.monthId}
                  className={`mileage-history-card ${entry.monthId === monthId ? "active" : ""}`}
                >
                  <button type="button" className="mileage-history-open" onClick={() => setMonthId(entry.monthId)}>
                    <span>{entry.monthLabel}</span>
                    <strong>{Number(entry.totalMiles || 0).toFixed(1)} miles</strong>
                    <small>{entry.lineCount} journey{entry.lineCount === 1 ? "" : "s"}</small>
                  </button>
                  <div className="mileage-history-journeys">
                    {(Array.isArray(entry.lines) ? entry.lines : []).map((line) => (
                      <div key={`${entry.monthId}-${line.id}`} className="mileage-history-journey">
                        <div>
                          <strong>{line.note || "No note"}</strong>
                          <small>{formatJobDate(line.date) || "-"}</small>
                          <span>{line.from || "-"} to {line.to || "-"}</span>
                        </div>
                        <span className="mileage-history-miles">{Number(line.miles || 0).toFixed(1)} miles</span>
                        <button
                          type="button"
                          className="text-button danger mileage-history-delete"
                          onClick={() => deleteMileageJourney(line.claimMonthId || entry.monthId, line.id, entry.monthLabel)}
                          disabled={deleting}
                        >
                          Delete
                        </button>
                      </div>
                    ))}
                  </div>
                </article>
              ))}
            </div>
          ) : (
            <div className="notifications-empty">No mileage submitted yet.</div>
          )}
        </section>
      </div>
    </div>
  );
}

function MileageAdminPage({ currentUser, onLogout, notifications }) {
  const initialMonth = useMemo(() => {
    const params = new URLSearchParams(typeof window !== "undefined" ? window.location.search : "");
    return params.get("month") || toMonthIdFromIso(getLocalTodayIso());
  }, []);
  const [monthId, setMonthId] = useState(initialMonth);
  const [overview, setOverview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState(null);

  async function loadMileageOverview(nextMonthId = monthId) {
    try {
      setLoading(true);
      const response = await fetch(`/api/mileage/admin?month=${encodeURIComponent(nextMonthId)}`);
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not load mileage overview.");
      setOverview(payload);
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not load mileage overview.", "error"));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadMileageOverview(monthId);
  }, [monthId]);

  useEffect(() => {
    if (!statusMessage) return undefined;
    const timer = window.setTimeout(() => setStatusMessage(null), 3500);
    return () => window.clearTimeout(timer);
  }, [statusMessage]);

  const users = Array.isArray(overview?.users) ? overview.users : [];
  const monthLabel = overview?.monthLabel || "Mileage overview";

  return (
    <div className="app-shell mileage-shell">
      <div className="page mileage-page mileage-admin-page">
        <MainNavBar currentUser={currentUser} active="mileage" onLogout={onLogout} notifications={notifications} />

        <section className="panel mileage-panel mileage-admin-panel">
          <div className="mileage-head">
            <div>
              <p className="eyebrow">Mileage</p>
              <h2>{monthLabel}</h2>
              <p>Admin overview of submitted journeys by each mileage user.</p>
            </div>
            <div className="mileage-month-tools">
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, -1))}>
                Previous month
              </button>
              <input
                type="month"
                value={monthId}
                onChange={(event) => setMonthId(event.target.value || toMonthIdFromIso(getLocalTodayIso()))}
              />
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, 1))}>
                Next month
              </button>
            </div>
          </div>

          {statusMessage ? <div className={`flash ${statusMessage.tone}`}>{statusMessage.text}</div> : null}

          <div className="mileage-admin-summary">
            <div>
              <span>Total miles</span>
              <strong>{Number(overview?.totalMiles || 0).toFixed(1)}</strong>
            </div>
            <div>
              <span>Journeys</span>
              <strong>{Number(overview?.lineCount || 0)}</strong>
            </div>
            <div>
              <span>Submitted</span>
              <strong>{Number(overview?.submittedUserCount || 0)} / {Number(overview?.userCount || 0)}</strong>
            </div>
          </div>

          {loading && !overview ? (
            <div className="notifications-empty">Loading mileage overview...</div>
          ) : (
            <div className="mileage-admin-users">
              {users.length ? (
                users.map((user) => (
                  <article key={user.userId || user.userName} className={`mileage-admin-user-card ${user.lineCount ? "" : "empty"}`}>
                    <div className="mileage-admin-user-head">
                      <div>
                        <h3>{user.userName}</h3>
                        <p>{user.lineCount} journey{user.lineCount === 1 ? "" : "s"} submitted</p>
                      </div>
                      <strong>{Number(user.totalMiles || 0).toFixed(1)} miles</strong>
                    </div>

                    {Array.isArray(user.journeys) && user.journeys.length ? (
                      <div className="mileage-admin-journeys">
                        {user.journeys.map((journey) => (
                          <div key={`${user.userId}-${journey.claimMonthId}-${journey.id}`} className="mileage-admin-journey">
                            <div className="mileage-admin-journey-main">
                              <strong>{journey.note || "No note"}</strong>
                              <span>{journey.from || "-"} to {journey.to || "-"}</span>
                            </div>
                            <time>{formatJobDate(journey.date) || "-"}</time>
                            <strong>{Number(journey.miles || 0).toFixed(1)} miles</strong>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="notifications-empty compact">No mileage submitted for this month.</div>
                    )}
                  </article>
                ))
              ) : (
                <div className="notifications-empty">No mileage users found.</div>
              )}
            </div>
          )}
        </section>
      </div>
    </div>
  );
}

function MileagePage(props) {
  return canEditMileage(props.currentUser) ? <MileageAdminPage {...props} /> : <MileageUserPage {...props} />;
}

function AttendancePage({
  currentUser,
  onLogout,
  notifications,
  attendanceData,
  loading,
  attendanceMonthId,
  setAttendanceMonthId,
  attendanceSavingKey,
  attendanceNoteSavingKey,
  attendanceFocusDate,
  onSaveAttendanceEntry,
  onSubmitAttendanceExplanation
}) {
  const adminMode = canEditAttendance(currentUser);
  const [drafts, setDrafts] = useState({});
  const [noteForm, setNoteForm] = useState(EMPTY_ATTENDANCE_NOTE_FORM);

  useEffect(() => {
    setDrafts({});
  }, [attendanceData?.monthId]);

  useEffect(() => {
    if (!attendanceFocusDate) return;
    setNoteForm((current) =>
      current.date === attendanceFocusDate
        ? current
        : {
            date: attendanceFocusDate,
            note: ""
          }
    );
  }, [attendanceFocusDate]);

  const rows = Array.isArray(attendanceData?.rows) ? attendanceData.rows : [];
  const staff = Array.isArray(attendanceData?.staff) ? attendanceData.staff : [];
  const missingEntries = Array.isArray(attendanceData?.missingEntries) ? attendanceData.missingEntries : [];
  const attendanceMonthLabel = attendanceData?.monthLabel || "Attendance";
  const focusedMissingEntry =
    missingEntries.find((entry) => entry.isoDate === noteForm.date) || missingEntries[0] || null;

  useEffect(() => {
    if (adminMode) return;
    if (noteForm.date) return;
    if (!missingEntries.length) return;
    setNoteForm({
      date: missingEntries[0].isoDate,
      note: missingEntries[0].employeeNote || ""
    });
  }, [adminMode, missingEntries, noteForm.date]);

  function getDraftValue(person, date, field, fallback = "") {
    const key = `${person}:${date}`;
    return drafts[key]?.[field] ?? fallback;
  }

  function setDraftValue(person, date, updates) {
    const key = `${person}:${date}`;
    setDrafts((current) => ({
      ...current,
      [key]: {
        ...(current[key] || {}),
        ...updates
      }
    }));
  }

  async function handleAttendanceBlur(cell) {
    const person = cell.person;
    const date = cell.isoDate;
    const key = `${person}:${date}`;
    const draft = drafts[key] || {};
    const clockIn = draft.clockIn ?? cell.clockIn ?? "";
    const clockOut = draft.clockOut ?? cell.clockOut ?? "";
    const adminNote = draft.adminNote ?? cell.adminNote ?? "";
    await onSaveAttendanceEntry({ person, date, clockIn, clockOut, adminNote });
  }

  async function submitEmployeeNote(event) {
    event.preventDefault();
    if (!noteForm.date || !noteForm.note.trim()) return;
    await onSubmitAttendanceExplanation({
      date: noteForm.date,
      note: noteForm.note.trim()
    });
    setNoteForm(EMPTY_ATTENDANCE_NOTE_FORM);
  }

  return (
    <div className="app-shell holidays-shell">
      <div className="page holidays-page attendance-page">
        <MainNavBar
          currentUser={currentUser}
          active="attendance"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel holidays-panel attendance-panel">
          <div className="holidays-toolbar attendance-toolbar">
            <div>
              <h2>Attendance</h2>
              <p>{attendanceMonthLabel}</p>
            </div>
            <div className="holidays-toolbar-actions">
              <button className="ghost-button" type="button" onClick={() => setAttendanceMonthId(shiftMonthId(attendanceMonthId, -1))}>
                Previous month
              </button>
              <button className="ghost-button" type="button" onClick={() => setAttendanceMonthId(toMonthIdFromIso(getLocalTodayIso()))}>
                Current month
              </button>
              <button className="ghost-button" type="button" onClick={() => setAttendanceMonthId(shiftMonthId(attendanceMonthId, 1))}>
                Next month
              </button>
            </div>
          </div>

          {loading ? <div className="board-loading">Loading attendance...</div> : null}

          {!loading && adminMode ? (
            <div className="attendance-grid-wrap">
              <table className="attendance-grid-table">
                <thead>
                  <tr>
                    <th className="attendance-date-head" rowSpan={2}>Date</th>
                    {staff.map((person) => (
                      <th
                        key={`staff-${person.person}`}
                        className="attendance-staff-head"
                        colSpan={2}
                        title={person.fullName || person.person}
                      >
                        <span>{person.code || person.fullName || person.person}</span>
                        <small>{person.fullName || person.person}</small>
                      </th>
                    ))}
                  </tr>
                  <tr>
                    {staff.map((person) => (
                      <>
                        <th key={`${person.person}-in`} className="attendance-sub-head">In</th>
                        <th key={`${person.person}-out`} className="attendance-sub-head">Out</th>
                      </>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row) => (
                    <tr key={row.isoDate} className={row.isToday ? "attendance-row-today" : ""}>
                      <th className="attendance-date-cell">
                        <span>{row.weekdayLabel}</span>
                        <strong>{row.dateLabel}</strong>
                        {row.isToday ? <em>Today</em> : null}
                      </th>
                      {row.cells.map((cell) => {
                        const missingClass = cell.hasMissingClock ? "attendance-cell-missing" : "";
                        const displayClass = getAttendanceDisplayClass(cell);
                        if (cell.displayLabel) {
                          return (
                            <td
                              key={`${row.isoDate}-${cell.person}`}
                              className={`attendance-merged-cell ${displayClass} ${missingClass}`.trim()}
                              colSpan={2}
                            >
                              <span className={cell.isHoliday ? "attendance-merged-holiday" : ""}>{cell.displayLabel}</span>
                            </td>
                          );
                        }
                          return (
                            <>
                              <td key={`${row.isoDate}-${cell.person}-in`} className={`attendance-value-cell ${missingClass}`}>
                                <input
                                  className="attendance-time-input"
                                  value={getDraftValue(cell.person, row.isoDate, "clockIn", cell.clockIn)}
                                  placeholder="--:--"
                                  onChange={(event) => setDraftValue(cell.person, row.isoDate, { clockIn: event.target.value })}
                                  onBlur={() => handleAttendanceBlur(cell)}
                                />
                                {cell.halfDayHolidayLabel ? (
                                  <span className="attendance-half-day-chip">{cell.halfDayHolidayLabel}</span>
                                ) : null}
                              </td>
                              <td key={`${row.isoDate}-${cell.person}-out`} className={`attendance-value-cell ${missingClass}`}>
                                <input
                                  className="attendance-time-input"
                                  value={getDraftValue(cell.person, row.isoDate, "clockOut", cell.clockOut)}
                                placeholder="--:--"
                                onChange={(event) => setDraftValue(cell.person, row.isoDate, { clockOut: event.target.value })}
                                onBlur={() => handleAttendanceBlur(cell)}
                              />
                            </td>
                          </>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}

          {!loading && !adminMode ? (
            <div className="attendance-self-service">
              <section className="attendance-self-panel">
                <h3>Your times</h3>
                <div className="attendance-self-list">
                  {rows.map((row) => {
                    const cell = row.cells[0];
                    if (!cell) return null;
                    return (
                      <article
                        key={row.isoDate}
                        className={`attendance-self-card ${row.isToday ? "is-today" : ""} ${cell.hasMissingClock ? "is-missing" : ""} ${noteForm.date === row.isoDate ? "is-focused" : ""}`}
                      >
                        <div className="attendance-self-card-head">
                          <div>
                            <strong>{row.dateLabel}</strong>
                            <span>{row.weekdayLabel}</span>
                          </div>
                          {cell.displayLabel ? <span className="attendance-self-status">{cell.displayLabel}</span> : null}
                           {!cell.displayLabel && cell.hasMissingClock ? <span className="attendance-self-status missing">Missing clocking</span> : null}
                           {!cell.displayLabel && cell.halfDayHolidayLabel ? (
                             <span className="attendance-self-status holiday">{cell.halfDayHolidayLabel}</span>
                           ) : null}
                          </div>
                        {cell.displayLabel ? null : (
                          <div className="attendance-self-times">
                            <span>In: <strong>{cell.clockIn || "--:--"}</strong></span>
                            <span>Out: <strong>{cell.clockOut || "--:--"}</strong></span>
                          </div>
                        )}
                        {!cell.displayLabel && cell.canExplain ? (
                          <button
                            type="button"
                            className="ghost-button attendance-note-button"
                            onClick={() => setNoteForm({ date: row.isoDate, note: cell.employeeNote || "" })}
                          >
                            {cell.employeeNote ? "Update note" : "Add note"}
                          </button>
                        ) : null}
                        {cell.employeeNote ? <p className="attendance-self-note">{cell.employeeNote}</p> : null}
                      </article>
                    );
                  })}
                </div>
              </section>

              <section className="attendance-self-panel">
                <h3>Missing clockings</h3>
                {missingEntries.length ? (
                  <>
                    <div className="attendance-missing-list">
                      {missingEntries.map((entry) => (
                        <button
                          key={`${entry.person}-${entry.isoDate}`}
                          type="button"
                          className={`attendance-missing-chip ${noteForm.date === entry.isoDate ? "active" : ""}`}
                          onClick={() => setNoteForm({ date: entry.isoDate, note: entry.employeeNote || "" })}
                        >
                          {entry.dateLabel}: {entry.clockIn ? "Missing out" : entry.clockOut ? "Missing in" : "Missing in/out"}
                        </button>
                      ))}
                    </div>
                    <form className="attendance-note-form" onSubmit={submitEmployeeNote}>
                      <label>
                        <span>Date</span>
                        <input type="text" value={focusedMissingEntry?.dateLabel || ""} readOnly />
                      </label>
                      <label>
                        <span>Explanation</span>
                        <textarea
                          rows={4}
                          value={noteForm.note}
                          onChange={(event) => setNoteForm((current) => ({ ...current, note: event.target.value }))}
                          placeholder="Explain the missing clocking so admin can check it."
                        />
                      </label>
                      <div className="attendance-note-actions">
                        <button
                          className="primary-button"
                          type="submit"
                          disabled={!noteForm.date || !noteForm.note.trim() || attendanceNoteSavingKey === noteForm.date}
                        >
                          {attendanceNoteSavingKey === noteForm.date ? "Saving..." : "Send note"}
                        </button>
                      </div>
                    </form>
                  </>
                ) : (
                  <div className="notifications-empty">No missing clockings for this month.</div>
                )}
              </section>
            </div>
          ) : null}
        </section>
      </div>
    </div>
  );
}

function VinylEstimatorPage({ currentUser, onLogout, notifications }) {
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

  function savePricingSettings() {
    const nextSettings = mergeVehiclePricingSettings(pricingDraftSettings);
    setPricingDraftSettings(nextSettings);
    setPricingSettingsByTemplate((current) => {
      const nextSettingsByTemplate = {
        ...current,
        [selectedTemplate.id]: nextSettings
      };
      saveVehiclePricingSettingsByTemplate(nextSettingsByTemplate);
      return nextSettingsByTemplate;
    });
  }

  function resetPricingDraft() {
    setPricingDraftSettings(pricingSettings);
  }

  function handlePricingSettingsToggle(event) {
    const isOpen = event.currentTarget.open;
    setPricingSettingsOpen(isOpen);
    setPricingDraftSettings(pricingSettings);
  }

  function restoreDefaultPricingSettings() {
    const defaultSettings = mergeVehiclePricingSettings(selectedTemplate.pricingSettings || {});
    setPricingDraftSettings(defaultSettings);
    setPricingSettingsByTemplate((current) => {
      const nextSettingsByTemplate = {
        ...current,
        [selectedTemplate.id]: defaultSettings
      };
      saveVehiclePricingSettingsByTemplate(nextSettingsByTemplate);
      return nextSettingsByTemplate;
    });
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

              {currentUser?.canManagePermissions ? (
                <details className="vinyl-pricing-settings" onToggle={handlePricingSettingsToggle}>
                  <summary>Pricing settings</summary>
                  <p className="vinyl-pricing-context">
                    Editing pricing for <strong>{selectedTemplate.sizeName}</strong> ({selectedTemplate.exampleName}).
                  </p>
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

export default function App() {
  const pathname = typeof window !== "undefined" ? window.location.pathname.toLowerCase() : "/";
  const search = typeof window !== "undefined" ? window.location.search : "";
  const isClientRoute = pathname.startsWith("/client");
  const isClientBoardRoute = pathname.startsWith("/client/board");
  const isInstallerRoute = pathname.startsWith("/installer");
  const isAttendanceRoute = pathname.startsWith("/attendance");
  const isHolidaysRoute = pathname.startsWith("/holidays");
  const isMileageRoute = pathname.startsWith("/mileage");
  const isVanEstimatorRoute = pathname.startsWith("/van-estimator");
  const isRamsLogicRoute = pathname.startsWith("/rams/logic");
  const isRamsRoute = pathname.startsWith("/rams");
  const isNotificationsRoute = pathname.startsWith("/notifications");
  const isBoardRoute = pathname.startsWith("/board");
  const [board, setBoard] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [holidays, setHolidays] = useState([]);
  const [holidayRequests, setHolidayRequests] = useState([]);
  const [approvedHolidayRequests, setApprovedHolidayRequests] = useState([]);
  const [holidayStaff, setHolidayStaff] = useState(HOLIDAY_STAFF);
  const [holidayAllowances, setHolidayAllowances] = useState([]);
  const [holidayEvents, setHolidayEvents] = useState([]);
  const [holidayAllowanceSavingKey, setHolidayAllowanceSavingKey] = useState("");
  const [holidayRequestOpen, setHolidayRequestOpen] = useState(false);
  const [holidayRequestForm, setHolidayRequestForm] = useState(EMPTY_HOLIDAY_REQUEST_FORM);
  const [holidayRequestSaving, setHolidayRequestSaving] = useState(false);
  const [holidayCancelOpen, setHolidayCancelOpen] = useState(false);
  const [holidayCancelForm, setHolidayCancelForm] = useState(EMPTY_HOLIDAY_CANCEL_FORM);
  const [holidayEventOpen, setHolidayEventOpen] = useState(false);
  const [holidayEventForm, setHolidayEventForm] = useState(EMPTY_HOLIDAY_EVENT_FORM);
  const [holidayEventSaving, setHolidayEventSaving] = useState(false);
  const [holidayYearStart, setHolidayYearStart] = useState(getCurrentHolidayYearStart());
  const [currentHolidayYearStart, setCurrentHolidayYearStart] = useState(getCurrentHolidayYearStart());
  const [attendanceData, setAttendanceData] = useState(null);
  const [attendanceMonthId, setAttendanceMonthId] = useState(toMonthIdFromIso(getLocalTodayIso()));
  const [attendanceLoading, setAttendanceLoading] = useState(false);
  const [attendanceSavingKey, setAttendanceSavingKey] = useState("");
  const [attendanceNoteSavingKey, setAttendanceNoteSavingKey] = useState("");
  const [form, setForm] = useState({ ...EMPTY_FORM, date: getLocalTodayIso() });
  const [editingId, setEditingId] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState(null);
  const [draggingJobId, setDraggingJobId] = useState("");
  const [duplicatingJobId, setDuplicatingJobId] = useState("");
  const [draggingHolidayId, setDraggingHolidayId] = useState("");
  const [dropDate, setDropDate] = useState("");
  const [activeHolidayDate, setActiveHolidayDate] = useState("");
  const [activeHolidayId, setActiveHolidayId] = useState("");
  const [jobModalDate, setJobModalDate] = useState("");
  const [activeClientJob, setActiveClientJob] = useState(null);
  const [clientCompletePrompt, setClientCompletePrompt] = useState(false);
  const [clientPhotoUploading, setClientPhotoUploading] = useState(false);
  const [clientExporting, setClientExporting] = useState(false);
  const [adminCompletePrompt, setAdminCompletePrompt] = useState(false);
  const [adminPhotoUploading, setAdminPhotoUploading] = useState(false);
  const [adminExporting, setAdminExporting] = useState(false);
  const [orderLookupOpen, setOrderLookupOpen] = useState(false);
  const [orderLookupQuery, setOrderLookupQuery] = useState("");
  const [orderLookupLoading, setOrderLookupLoading] = useState(false);
  const [orderLookupResults, setOrderLookupResults] = useState([]);
  const [orderLookupError, setOrderLookupError] = useState("");
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [loginUsers, setLoginUsers] = useState([]);
  const [loginDisplayName, setLoginDisplayName] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [permissionSavingKey, setPermissionSavingKey] = useState("");
  const [notifications, setNotifications] = useState([]);
  const [previousMonthDepth, setPreviousMonthDepth] = useState(0);
  const [futureMonthDepth, setFutureMonthDepth] = useState(0);
  const boardNotificationJobId = useMemo(() => {
    const params = new URLSearchParams(search);
    return params.get("job") || "";
  }, [search]);
  const attendanceNotificationDate = useMemo(() => {
    const params = new URLSearchParams(search);
    return params.get("date") || "";
  }, [search]);
  const dragPreviewRef = useRef(null);
  const transparentDragImageRef = useRef(null);
  const dragPositionRef = useRef({ x: 0, y: 0 });
  const clientPhotoInputRef = useRef(null);
  const adminPhotoInputRef = useRef(null);
  const openedNotificationJobIdRef = useRef("");
  const boardEditable = canEditBoard(currentUser);
  const installerEditable = canEditInstaller(currentUser);
  const attendanceEditable = canEditAttendance(currentUser);
  const hostShellMode = usesHostShell(currentUser);
  const isClientMode = currentUser ? !boardEditable : false;
  const showInstallerDirectory = Boolean(currentUser && canAccessInstaller(currentUser) && isInstallerRoute);
  const showAttendance = Boolean(currentUser && canAccessAttendance(currentUser) && isAttendanceRoute);
  const showHolidays = Boolean(currentUser && canAccessHolidays(currentUser) && isHolidaysRoute);
  const showMileage = Boolean(currentUser && canAccessMileage(currentUser) && isMileageRoute);
  const showVanEstimator = Boolean(currentUser && canAccessVanEstimator(currentUser) && isVanEstimatorRoute);
  const showRamsLogic = Boolean(currentUser && canAccessRams(currentUser) && isRamsLogicRoute);
  const showRams = Boolean(currentUser && canAccessRams(currentUser) && isRamsRoute && !isRamsLogicRoute);
  const showNotifications = Boolean(currentUser && isNotificationsRoute);
  const showBoard = Boolean(
    currentUser &&
      canAccessBoard(currentUser) &&
      ((boardEditable && isBoardRoute) || (!boardEditable && isClientBoardRoute))
  );
  const showHostLanding = Boolean(currentUser && hostShellMode && !isInstallerRoute && !isBoardRoute && !isClientBoardRoute && !isAttendanceRoute && !isHolidaysRoute && !isMileageRoute && !isVanEstimatorRoute && !isRamsRoute && !isNotificationsRoute);
  const showClientLanding = Boolean(currentUser && !hostShellMode && (canAccessBoard(currentUser) || canAccessAttendance(currentUser) || canAccessHolidays(currentUser) || canAccessMileage(currentUser) || canAccessVanEstimator(currentUser)) && !isClientBoardRoute && !isAttendanceRoute && !isHolidaysRoute && !isMileageRoute && !isVanEstimatorRoute && !isRamsRoute && !isNotificationsRoute);
  const activeAdminJob = useMemo(() => {
    if (!editingId) return null;
    return jobs.find((job) => String(job.id || "") === String(editingId)) || null;
  }, [editingId, jobs]);

  const todayIso = board?.today || getLocalTodayIso();
  const rollingStartIso = useMemo(() => {
    const today = parseIsoDate(todayIso);
    return today ? toIsoDate(addDays(today, -7)) : "";
  }, [todayIso]);
  const rollingEndIso = useMemo(() => {
    const today = parseIsoDate(todayIso);
    return today ? toIsoDate(addDays(today, 21)) : "";
  }, [todayIso]);
  const boardRange = useMemo(() => {
    const today = parseIsoDate(todayIso);
    if (!today) {
      return { startIso: "", endIso: "" };
    }

    const currentMonthStart = getStartOfMonth(today);
    const start =
      previousMonthDepth > 0
        ? getStartOfMonth(addMonths(currentMonthStart, -previousMonthDepth))
        : parseIsoDate(rollingStartIso);
    const end =
      futureMonthDepth > 0
        ? getEndOfMonth(addMonths(currentMonthStart, futureMonthDepth))
        : parseIsoDate(rollingEndIso);

    return {
      startIso: start ? toIsoDate(start) : "",
      endIso: end ? toIsoDate(end) : ""
    };
  }, [futureMonthDepth, previousMonthDepth, rollingEndIso, rollingStartIso, todayIso]);

  function resetBoardWindow() {
    setPreviousMonthDepth(0);
    setFutureMonthDepth(0);
  }

  useEffect(() => {
    let active = true;

    async function loadAuth() {
      try {
        const meResponse = await fetch("/api/auth/me");
        const mePayload = meResponse.ok ? await meResponse.json() : null;
        let usersPayload = [];

        if (mePayload?.user) {
          const usersResponse = await fetch("/api/auth/users");
          usersPayload = usersResponse.ok ? await usersResponse.json() : [];
        }

        if (!active) return;

        setLoginUsers(Array.isArray(usersPayload) ? usersPayload : []);
        if (mePayload?.user) {
          setCurrentUser(mePayload.user);
        }
      } catch (error) {
        console.error(error);
      } finally {
        if (active) setAuthChecked(true);
      }
    }

    loadAuth();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!currentUser) {
      setNotifications([]);
      return undefined;
    }

    let active = true;

    async function loadNotifications() {
      try {
        const response = await fetch("/api/notifications");
        if (!response.ok) {
          throw new Error("Could not load notifications.");
        }
        const payload = await response.json();
        if (!active) return;
        setNotifications(Array.isArray(payload) ? payload : []);
      } catch (error) {
        console.error(error);
      }
    }

    loadNotifications();
    return () => {
      active = false;
    };
  }, [currentUser]);

  useEffect(() => {
    if (!currentUser || !showBoard) return undefined;
    let active = true;

    async function loadBoard() {
      try {
        setLoading(true);
        const [boardResponse, jobsResponse, holidaysResponse] = await Promise.all([
          fetch(buildBoardUrl(boardRange.startIso, boardRange.endIso)),
          fetch("/api/jobs"),
          fetch("/api/holidays")
        ]);
        if (!boardResponse.ok || !jobsResponse.ok || !holidaysResponse.ok) {
          throw new Error("Could not load the installation board.");
        }

        const [boardData, jobsData, holidaysData] = await Promise.all([
          boardResponse.json(),
          jobsResponse.json(),
          holidaysResponse.json()
        ]);
        if (!active) return;
        setBoard(boardData);
        setJobs(Array.isArray(jobsData) ? jobsData : []);
        setHolidays(Array.isArray(holidaysData) ? holidaysData : []);
      } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage("Could not load the shared board.", "error"));
      } finally {
        if (active) setLoading(false);
      }
    }

    loadBoard();
    return () => {
      active = false;
    };
  }, [currentUser, showBoard, boardRange.endIso, boardRange.startIso]);

  useEffect(() => {
    if (!showBoard || !boardNotificationJobId || !Array.isArray(jobs) || !jobs.length) return;
    if (openedNotificationJobIdRef.current === String(boardNotificationJobId)) return;

    const matchedJob = jobs.find((job) => String(job.id || "") === String(boardNotificationJobId));
    if (!matchedJob) return;

    openedNotificationJobIdRef.current = String(boardNotificationJobId);
    if (isClientMode) {
      setActiveClientJob(matchedJob);
    } else {
      editJob(matchedJob);
    }

    const params = new URLSearchParams(window.location.search);
    params.delete("job");
    const nextSearch = params.toString();
    const nextUrl = `${window.location.pathname}${nextSearch ? `?${nextSearch}` : ""}${window.location.hash || ""}`;
    window.history.replaceState({}, "", nextUrl);
  }, [boardNotificationJobId, isClientMode, jobs, showBoard]);

  useEffect(() => {
    setClientCompletePrompt(false);
    setClientPhotoUploading(false);
    setClientExporting(false);
    if (clientPhotoInputRef.current) {
      clientPhotoInputRef.current.value = "";
    }
  }, [activeClientJob?.id]);

  useEffect(() => {
    setAdminCompletePrompt(false);
    setAdminPhotoUploading(false);
    setAdminExporting(false);
    if (adminPhotoInputRef.current) {
      adminPhotoInputRef.current.value = "";
    }
  }, [activeAdminJob?.id, jobModalDate]);

  useEffect(() => {
    if (!currentUser || !showHolidays) return undefined;
    const stream = new EventSource("/api/events");

    async function handleUpdate() {
      try {
        await refreshHolidayData();
      } catch (error) {
        console.error(error);
      }
    }

    stream.addEventListener("board-updated", handleUpdate);
    stream.onerror = () => {
      stream.close();
    };

    return () => {
      stream.removeEventListener("board-updated", handleUpdate);
      stream.close();
    };
  }, [currentUser, showHolidays]);

  useEffect(() => {
    if (!currentUser || !showHolidays) return undefined;
    let active = true;

    async function loadHolidayData() {
      try {
        setLoading(true);
        const response = await fetch(`/api/holiday-requests?yearStart=${encodeURIComponent(holidayYearStart)}`);
        if (!response.ok) {
          throw new Error("Could not load holiday calendar.");
        }

          const payload = await response.json();
          if (!active) return;
          setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
          setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
          setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
          setHolidayStaff(normalizeHolidayStaffEntries(payload.holidayStaff));
          setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
          setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
          setCurrentHolidayYearStart(Number(payload.currentHolidayYearStart || getCurrentHolidayYearStart()));
        } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage("Could not load holiday calendar.", "error"));
      } finally {
        if (active) setLoading(false);
      }
    }

    loadHolidayData();
    return () => {
      active = false;
    };
  }, [currentUser, showHolidays, holidayYearStart]);

  useEffect(() => {
    if (!showAttendance) return;
    if (!attendanceNotificationDate) return;
    const focusMonthId = toMonthIdFromIso(attendanceNotificationDate);
    if (focusMonthId) {
      setAttendanceMonthId(focusMonthId);
    }
  }, [attendanceNotificationDate, showAttendance]);

  useEffect(() => {
    if (!currentUser || !showAttendance) return undefined;
    let active = true;

    async function loadAttendance() {
      try {
        setAttendanceLoading(true);
        const response = await fetch(`/api/attendance?month=${encodeURIComponent(attendanceMonthId)}`);
        if (!response.ok) {
          throw new Error("Could not load attendance.");
        }
        const payload = await response.json();
        if (!active) return;
        setAttendanceData(payload);
      } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage(error.message || "Could not load attendance.", "error"));
      } finally {
        if (active) setAttendanceLoading(false);
      }
    }

    loadAttendance();
    return () => {
      active = false;
    };
  }, [attendanceMonthId, currentUser, showAttendance]);

  useEffect(() => {
    if (!currentUser) return;
    const nextHomePath = getHomePathForUser(currentUser);
    const nextBoardPath = getBoardPathForUser(currentUser);

    if (isHolidaysRoute && !canAccessHolidays(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isAttendanceRoute && !canAccessAttendance(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isMileageRoute && !canAccessMileage(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isVanEstimatorRoute && !canAccessVanEstimator(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isRamsRoute && !canAccessRams(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isNotificationsRoute && !currentUser) {
      window.location.replace("/");
      return;
    }

    if (isInstallerRoute && !canAccessInstaller(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if ((isBoardRoute || isClientBoardRoute) && !canAccessBoard(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (hostShellMode && isClientRoute && !isClientBoardRoute) {
      window.location.replace(nextHomePath);
      return;
    }

    if (!hostShellMode && !isClientRoute && !isHolidaysRoute && !isAttendanceRoute && !isMileageRoute && !isVanEstimatorRoute && !isRamsRoute && !isNotificationsRoute) {
      window.location.replace(nextHomePath);
      return;
    }

    if ((isBoardRoute || isClientBoardRoute) && nextBoardPath !== window.location.pathname) {
      window.location.replace(nextBoardPath);
    }
  }, [currentUser, isClientRoute, isClientBoardRoute, isInstallerRoute, isBoardRoute, isAttendanceRoute, isHolidaysRoute, isMileageRoute, isVanEstimatorRoute, isRamsRoute, isRamsLogicRoute, isNotificationsRoute, hostShellMode]);

  useEffect(() => {
    if (!currentUser || !showBoard) return undefined;
    const stream = new EventSource("/api/events");

    async function handleUpdate() {
      try {
        const response = await fetch(buildBoardUrl(boardRange.startIso, boardRange.endIso));
        if (!response.ok) return;
        const nextBoard = await response.json();
        setBoard(nextBoard);
      } catch (error) {
        console.error(error);
      }
    }

    stream.addEventListener("board-updated", handleUpdate);
    stream.onerror = () => {
      stream.close();
      window.setTimeout(() => window.location.reload(), 3000);
    };

    return () => {
      stream.removeEventListener("board-updated", handleUpdate);
      stream.close();
    };
  }, [currentUser, showBoard, boardRange.endIso, boardRange.startIso]);

  useEffect(() => {
    if (!currentUser || !showAttendance) return undefined;
    const stream = new EventSource("/api/events");

    async function handleUpdate() {
      try {
        const response = await fetch(`/api/attendance?month=${encodeURIComponent(attendanceMonthId)}`);
        if (!response.ok) return;
        const payload = await response.json();
        setAttendanceData(payload);
      } catch (error) {
        console.error(error);
      }
    }

    stream.addEventListener("attendance-updated", handleUpdate);
    stream.onerror = () => {
      stream.close();
    };

    return () => {
      stream.removeEventListener("attendance-updated", handleUpdate);
      stream.close();
    };
  }, [attendanceMonthId, currentUser, showAttendance]);

  useEffect(() => {
    if (!message) return undefined;
    const timer = window.setTimeout(() => setMessage(null), 3000);
    return () => window.clearTimeout(timer);
  }, [message]);

  useEffect(() => {
    function handleWindowDragOver(event) {
      if (!dragPreviewRef.current) return;
      const deltaX = event.clientX - dragPositionRef.current.x;
      const deltaY = event.clientY - dragPositionRef.current.y;
      dragPositionRef.current = { x: event.clientX, y: event.clientY };

      const tilt = Math.max(-12, Math.min(12, deltaX * 0.6));
      const lift = Math.max(-8, Math.min(8, -deltaY * 0.25));
      dragPreviewRef.current.style.left = `${event.clientX + 18}px`;
      dragPreviewRef.current.style.top = `${event.clientY + 18}px`;
      dragPreviewRef.current.style.transform = `rotate(${tilt}deg) translateY(${lift}px)`;
    }

    function handleWindowDrop() {
      clearDragPreview();
    }

    window.addEventListener("dragover", handleWindowDragOver);
    window.addEventListener("drop", handleWindowDrop);

    return () => {
      window.removeEventListener("dragover", handleWindowDragOver);
      window.removeEventListener("drop", handleWindowDrop);
    };
  }, []);

  const upcomingJobs = useMemo(() => {
    const today = board?.today || getLocalTodayIso();
    return [...jobs]
      .filter((job) => job.date >= today)
      .sort((left, right) => {
        if (left.date !== right.date) return left.date.localeCompare(right.date);
        return left.customerName.localeCompare(right.customerName);
      })
      .slice(0, 8);
  }, [board?.today, jobs]);

  const jobsById = useMemo(() => {
    return new Map(jobs.map((job) => [job.id, job]));
  }, [jobs]);

  const holidaysById = useMemo(() => {
    return new Map(holidays.map((holiday) => [holiday.id, holiday]));
  }, [holidays]);

  const holidayYearLabel = useMemo(() => getHolidayYearLabel(holidayYearStart), [holidayYearStart]);
  const holidayRows = useMemo(
    () => buildHolidayYearRows(holidays, holidayYearStart, holidayEvents),
    [holidayEvents, holidayYearStart, holidays]
  );

  const backdropPointerStartedRef = useRef(false);

  function resetForm(nextDate = board?.today || getLocalTodayIso()) {
    setEditingId("");
    setForm({ ...EMPTY_FORM, date: nextDate });
    setJobModalDate("");
    setOrderLookupOpen(false);
    setOrderLookupQuery("");
    setOrderLookupResults([]);
    setOrderLookupError("");
  }

  function preserveScrollPosition() {
    const currentScrollY = window.scrollY;
    window.requestAnimationFrame(() => {
      window.scrollTo({ top: currentScrollY, behavior: "auto" });
    });
  }

  function openJobModal(nextDate, nextForm, nextEditingId = "") {
    setEditingId(nextEditingId);
    setForm(nextForm);
    setJobModalDate(nextDate);
    preserveScrollPosition();
  }

  function handleBackdropPointerDown(event) {
    backdropPointerStartedRef.current = event.target === event.currentTarget;
  }

  function handleBackdropClick(event, onClose) {
    const shouldClose = backdropPointerStartedRef.current && event.target === event.currentTarget;
    backdropPointerStartedRef.current = false;
    if (shouldClose) {
      onClose();
    }
  }

  async function handleLogin(event) {
    event.preventDefault();
    try {
      setLoginLoading(true);
      setLoginError("");
      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          displayName: loginDisplayName,
          password: loginPassword
        })
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not sign in.");
      }

      setCurrentUser(payload.user);
      setLoginPassword("");
      window.location.replace(getHomePathForUser(payload.user));
    } catch (error) {
      console.error(error);
      setLoginError(error.message || "Could not sign in.");
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleLogout() {
    try {
      await fetch("/api/auth/logout", { method: "POST" });
    } catch (error) {
      console.error(error);
    } finally {
      setCurrentUser(null);
      setBoard(null);
      setJobs([]);
      setHolidays([]);
      setHolidayRequests([]);
      setApprovedHolidayRequests([]);
      setLoginPassword("");
      setLoginError("");
      window.location.replace(isClientRoute ? "/client" : "/");
    }
  }

  async function handlePermissionChange(userId, appKey, value) {
    const targetUser = loginUsers.find((entry) => entry.id === userId);
    if (!targetUser || !currentUser?.canManagePermissions) return;

    const nextPermissions = {
      board: getPermissionForApp(targetUser, "board"),
      installer: getPermissionForApp(targetUser, "installer"),
      holidays: getPermissionForApp(targetUser, "holidays"),
      attendance: getPermissionForApp(targetUser, "attendance"),
      mileage: getPermissionForApp(targetUser, "mileage"),
      vanEstimator: getPermissionForApp(targetUser, "vanEstimator"),
      [appKey]: value
    };

    setPermissionSavingKey(`${userId}:${appKey}`);

    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(userId)}/permissions`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(nextPermissions)
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not update permissions.");
      }

      setLoginUsers((current) =>
        current.map((entry) => (entry.id === userId ? { ...entry, ...payload.user } : entry))
      );
      setMessage(createMessage(`Updated ${targetUser.displayName}'s permissions.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update permissions.", "error"));
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleUpdateAttendanceProfile(userId, attendanceProfile) {
    const targetUser = loginUsers.find((entry) => entry.id === userId);
    if (!targetUser || !currentUser?.canManagePermissions) return;

    setPermissionSavingKey(`${userId}:attendance-profile`);

    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(userId)}/attendance-profile`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(attendanceProfile)
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not update attendance settings.");
      }

      setLoginUsers((current) =>
        current.map((entry) => (entry.id === userId ? { ...entry, ...payload.user } : entry))
      );
      if (currentUser?.id === userId) {
        setCurrentUser((existing) => ({ ...existing, ...payload.user }));
      }
      if (showAttendance) {
        const attendanceResponse = await fetch(`/api/attendance?month=${encodeURIComponent(attendanceMonthId)}`);
        if (attendanceResponse.ok) {
          const attendancePayload = await attendanceResponse.json();
          setAttendanceData(attendancePayload);
        }
      }
      setMessage(createMessage(`Updated ${targetUser.displayName}'s attendance settings.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update attendance settings.", "error"));
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleCreateUser({ displayName, role, password }) {
    setPermissionSavingKey("create-user");
    try {
      const response = await fetch("/api/auth/users", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ displayName, role, password })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not create user.");
      }

      setLoginUsers((current) => [...current, payload.user].sort((left, right) => left.displayName.localeCompare(right.displayName)));
      setMessage(createMessage(`Added ${payload.user.displayName}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not create user.", "error"));
      throw error;
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleResetUserPassword(userId, password) {
    setPermissionSavingKey(`${userId}:password`);
    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(userId)}/password`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not update password.");
      }

      setLoginUsers((current) =>
        current.map((entry) => (entry.id === userId ? { ...entry, ...payload.user } : entry))
      );
      setMessage(createMessage(`Updated ${payload.user.displayName}'s password.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update password.", "error"));
      throw error;
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleDeleteUser(user) {
    if (!window.confirm(`Delete ${user.displayName}?`)) return;
    setPermissionSavingKey(`${user.id}:delete`);
    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(user.id)}`, {
        method: "DELETE"
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not delete user.");
      }

      setLoginUsers((current) => current.filter((entry) => entry.id !== user.id));
      setMessage(createMessage(`Deleted ${user.displayName}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not delete user.", "error"));
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function markNotificationRead(notificationId) {
    setNotifications((current) =>
      current.map((entry) =>
        entry.id === notificationId
          ? {
              ...entry,
              read: true
            }
          : entry
      )
    );

    try {
      const response = await fetch(`/api/notifications/${encodeURIComponent(notificationId)}/read`, {
        method: "PATCH"
      });
      if (!response.ok) {
        throw new Error("Could not update notification.");
      }
      const payload = await response.json();
      setNotifications(Array.isArray(payload) ? payload : []);
    } catch (error) {
      await refreshNotifications();
      throw error;
    }
  }

  async function markAllNotificationsRead() {
    try {
      setNotifications((current) => current.map((entry) => ({ ...entry, read: true })));
      const response = await fetch("/api/notifications/read-all", { method: "POST" });
      if (!response.ok) {
        throw new Error("Could not update notifications.");
      }
      const payload = await response.json();
      setNotifications(Array.isArray(payload) ? payload : []);
    } catch (error) {
      await refreshNotifications();
      console.error(error);
      setMessage(createMessage(error.message || "Could not update notifications.", "error"));
    }
  }

  async function openNotification(notification) {
    try {
      if (!notification?.read) {
        await markNotificationRead(notification.id);
      }
    } catch (error) {
      console.error(error);
    }

    if (notification?.link) {
      window.location.assign(notification.link);
    }
  }

  function editJob(job) {
    openJobModal(job.date || "Unscheduled", {
      id: job.id,
      date: job.date,
      orderReference: job.orderReference || "",
      customerName: job.customerName || "",
      description: job.description || "",
      contact: job.contact || "",
      number: job.number || "",
      address: job.address || "",
      installers: Array.isArray(job.installers)
        ? job.installers
        : typeof job.installers === "string" && job.installers.trim()
          ? job.installers.split(/[,/]+/).map((item) => item.trim()).filter(Boolean)
          : [],
      customInstaller: job.customInstaller || "",
      jobType: job.jobType || "Install",
      customJobType: job.customJobType || "",
      isPlaceholder: Boolean(job.isPlaceholder),
      notes: job.notes || ""
    }, job.id);
  }

  async function refreshData() {
    const [jobsResponse, holidaysResponse] = await Promise.all([fetch("/api/jobs"), fetch("/api/holidays")]);
    if (!jobsResponse.ok || !holidaysResponse.ok) throw new Error("Could not refresh data.");
    const [nextJobs, nextHolidays] = await Promise.all([jobsResponse.json(), holidaysResponse.json()]);
    setJobs(Array.isArray(nextJobs) ? nextJobs : []);
    setHolidays(Array.isArray(nextHolidays) ? nextHolidays : []);
  }

  async function refreshHolidayData() {
    const response = await fetch(`/api/holiday-requests?yearStart=${encodeURIComponent(holidayYearStart)}`);
    if (!response.ok) throw new Error("Could not refresh holiday calendar.");
      const payload = await response.json();
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidayStaff(normalizeHolidayStaffEntries(payload.holidayStaff));
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayYearStart(Number(payload.holidayYearStart || holidayYearStart));
      setCurrentHolidayYearStart(Number(payload.currentHolidayYearStart || getCurrentHolidayYearStart()));
    }

  async function refreshNotifications() {
    const response = await fetch("/api/notifications");
    if (!response.ok) throw new Error("Could not refresh notifications.");
    const payload = await response.json();
    setNotifications(Array.isArray(payload) ? payload : []);
  }

  async function saveAttendanceEntry({ person, date, clockIn, clockOut, adminNote = "" }) {
    setAttendanceSavingKey(`${person}:${date}`);
    try {
      const response = await fetch("/api/attendance/entries", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ person, date, clockIn, clockOut, adminNote })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not save attendance.");
      }
      setAttendanceData(payload);
      setMessage(createMessage(`Updated attendance for ${person} on ${formatJobDate(date)}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not save attendance.", "error"));
    } finally {
      setAttendanceSavingKey("");
    }
  }

  async function submitAttendanceExplanation({ date, note }) {
    setAttendanceNoteSavingKey(date);
    try {
      const response = await fetch("/api/attendance/explanations", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ date, note })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not send attendance note.");
      }
      setAttendanceData(payload);
      await refreshNotifications();
      setMessage(createMessage("Attendance note sent.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not send attendance note.", "error"));
    } finally {
      setAttendanceNoteSavingKey("");
    }
  }

  async function searchCoreBridgeOrders(searchTerm = orderLookupQuery) {
    if (isClientMode) return;

    try {
      setOrderLookupLoading(true);
      setOrderLookupError("");
      const query = String(searchTerm || "").trim();
      const params = new URLSearchParams();
      if (query) params.set("q", query);
      const url = params.toString() ? `/api/corebridge/orders?${params.toString()}` : "/api/corebridge/orders";
      const response = await fetch(url);
      const raw = await response.text();
      let payload = {};

      try {
        payload = raw ? JSON.parse(raw) : {};
      } catch (error) {
        throw new Error("CoreBridge returned an invalid response.");
      }

      if (!response.ok) {
        throw new Error(payload.detail || payload.error || "Could not load CoreBridge orders.");
      }

      setOrderLookupResults(Array.isArray(payload.orders) ? payload.orders : []);
    } catch (error) {
      console.error(error);
      setOrderLookupResults([]);
      setOrderLookupError(error.message || "Could not load CoreBridge orders.");
    } finally {
      setOrderLookupLoading(false);
    }
  }

  async function submitHolidayRequest() {
    const person = canEditHolidays(currentUser)
      ? holidayRequestForm.person
      : getHolidayStaffPersonForUser(currentUser);

    if (!person || !holidayRequestForm.startDate || !holidayRequestForm.endDate) {
      setMessage(createMessage("Choose a person plus start and end dates.", "error"));
      return;
    }

    setHolidayRequestSaving(true);

    try {
      const response = await fetch("/api/holiday-requests", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          person,
          holidayYearStart,
          startDate: holidayRequestForm.startDate,
          endDate: holidayRequestForm.endDate,
          duration: holidayRequestForm.duration,
          notes: holidayRequestForm.notes
        })
      });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Could not send holiday request.");
        setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
        setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
        setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
        setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
        setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
        setHolidayRequestForm(EMPTY_HOLIDAY_REQUEST_FORM);
        setHolidayRequestOpen(false);
        await refreshNotifications();
        setMessage(createMessage("Holiday request sent.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not send holiday request.", "error"));
    } finally {
      setHolidayRequestSaving(false);
    }
  }

  async function reviewHolidayRequest(requestId, status) {
    try {
      const response = await fetch(`/api/holiday-requests/${requestId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ status })
      });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Could not update holiday request.");
        setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
        setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
        setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
        setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
        setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
        await refreshNotifications();
        setMessage(createMessage(`Holiday request ${status}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update holiday request.", "error"));
    }
  }

  async function saveHolidayAllowance(person, updates) {
    if (!canEditHolidays(currentUser)) return;
    const existing = holidayAllowances.find((entry) => entry.person === person);
    const savingField = Object.keys(updates)[0] || "";
    setHolidayAllowanceSavingKey(`${person}:${savingField}`);

    try {
      const response = await fetch("/api/holiday-allowances", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...existing,
          person,
          yearStart: holidayYearStart,
          ...updates
        })
      });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Could not save holiday allowance.");
        setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
        setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
        setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
        setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
        setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
        setMessage(createMessage(`Updated ${person}'s holiday allowance.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not save holiday allowance.", "error"));
    } finally {
      setHolidayAllowanceSavingKey("");
    }
  }

  function changeHolidayAllowanceDraft(person, updates) {
    setHolidayAllowances((current) =>
      current.map((entry) =>
        entry.person === person
          ? {
              ...entry,
              ...updates
            }
          : entry
      )
    );
  }

  async function openOrderLookup() {
    setOrderLookupOpen(true);
    setOrderLookupError("");
  }

  async function applyCoreBridgeOrder(order) {
    let resolvedOrder = order;

    try {
      if (order?.id) {
        setOrderLookupLoading(true);
        setOrderLookupError("");
        const url = `/api/corebridge/orders/${encodeURIComponent(order.id)}`;
        const response = await fetch(url);
        const raw = await response.text();
        const payload = raw ? JSON.parse(raw) : {};
        if (!response.ok) {
          throw new Error(payload.detail || payload.error || "Could not load CoreBridge order detail.");
        }
        resolvedOrder = payload;
      }
    } catch (error) {
      console.error(error);
      setOrderLookupError(error.message || "Could not load CoreBridge order detail.");
      return;
    } finally {
      setOrderLookupLoading(false);
    }

    setForm((current) => ({
      ...current,
      orderReference: resolvedOrder.orderReference ?? "",
      customerName: resolvedOrder.customerName ?? "",
      description: resolvedOrder.description ?? "",
      contact: resolvedOrder.contact ?? "",
      number: resolvedOrder.number ?? "",
      address: resolvedOrder.address ?? "",
      notes: resolvedOrder.notes ?? ""
    }));
    setOrderLookupOpen(false);
    setMessage(createMessage("Order details copied into the job form.", "success"));
  }

  function applyBoardPayloadToState(payload, fallbackJobId = "") {
    if (payload?.board) setBoard(payload.board);
    if (Array.isArray(payload?.jobs)) setJobs(payload.jobs);
    if (Array.isArray(payload?.holidays)) setHolidays(payload.holidays);

    if (payload?.job) {
      setActiveClientJob(payload.job);
      return;
    }

    if (fallbackJobId && Array.isArray(payload?.jobs)) {
      const matched = payload.jobs.find((job) => String(job.id || "") === String(fallbackJobId));
      if (matched) {
        setActiveClientJob(matched);
      }
    }
  }

  async function cancelHolidayRequest() {
    const currentPerson = getHolidayStaffPersonForUser(currentUser);
    const currentPersonKey = getHolidayStaffIdentityKey(currentPerson);
    const currentUserId = String(currentUser?.id || "");
    const cancellableRequests = approvedHolidayRequests.filter((request) => {
      const requestStatus = String(request.status || "").trim().toLowerCase();
      const requestAction = String(request.action || "book").trim().toLowerCase();
      const sameUser =
        (currentUserId && String(request.requestedByUserId || "") === currentUserId) ||
        getHolidayStaffIdentityKey(request.person) === currentPersonKey;
      return sameUser && requestStatus === "approved" && requestAction === "book";
    });
    const targetHolidayRequest = cancellableRequests.find(
      (request) => String(request.id || "") === String(holidayCancelForm.requestId || "")
    );
    if (!targetHolidayRequest) {
      setMessage(createMessage("Choose an approved holiday request to cancel.", "error"));
      return;
    }

    setHolidayRequestSaving(true);
    try {
      const response = await fetch("/api/holiday-requests", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          person: targetHolidayRequest.person,
          holidayYearStart,
          startDate: targetHolidayRequest.startDate,
          endDate: targetHolidayRequest.endDate,
          duration: targetHolidayRequest.duration,
          notes: holidayCancelForm.notes,
          action: "cancel",
          targetRequestId: targetHolidayRequest.id
        })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not send cancellation request.");
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayCancelForm(EMPTY_HOLIDAY_CANCEL_FORM);
      setHolidayCancelOpen(false);
      await refreshNotifications();
      setMessage(createMessage("Holiday cancellation request sent.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not send cancellation request.", "error"));
    } finally {
      setHolidayRequestSaving(false);
    }
  }

  async function completeJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/complete`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not mark the job as complete.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function snaggingJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/snagging`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not mark the job as snagging.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function clearSnaggingJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/unsnagging`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not remove the snagging tag.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function undoCompleteJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/uncomplete`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not undo the completed status.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function uploadJobPhotos(jobId, files) {
    for (const file of files) {
      const prepared = await compressPhotoForUpload(file);
      const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/photos`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(prepared)
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || `Could not upload ${file.name}.`);
      }
      applyBoardPayloadToState(payload, jobId);
    }
  }

  async function markClientJobComplete(job, uploadFiles = []) {
    if (!job?.id) return;
    const files = Array.from(uploadFiles || []);
    if (files.length) {
      setClientPhotoUploading(true);
    }

    try {
      if (!job.isCompleted) {
        await completeJob(job.id);
      }
      if (files.length) {
        await uploadJobPhotos(job.id, files);
      }
      setClientCompletePrompt(false);
      setMessage(
        createMessage(
          files.length ? "Job marked complete and photos uploaded." : "Job marked complete.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not complete the job.", "error"));
    } finally {
      setClientPhotoUploading(false);
      if (clientPhotoInputRef.current) {
        clientPhotoInputRef.current.value = "";
      }
    }
  }

  async function markAdminJobComplete(job, uploadFiles = []) {
    if (!job?.id) return;
    const files = Array.from(uploadFiles || []);
    if (files.length) {
      setAdminPhotoUploading(true);
    }

    try {
      if (!job.isCompleted) {
        await completeJob(job.id);
      }
      if (files.length) {
        await uploadJobPhotos(job.id, files);
      }
      setAdminCompletePrompt(false);
      setMessage(
        createMessage(
          files.length ? "Job marked complete and photos uploaded." : "Job marked complete.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not complete the job.", "error"));
    } finally {
      setAdminPhotoUploading(false);
      if (adminPhotoInputRef.current) {
        adminPhotoInputRef.current.value = "";
      }
    }
  }

  async function markAdminJobSnagging(job) {
    if (!job?.id) return;
    try {
      await snaggingJob(job.id);
      setAdminCompletePrompt(false);
      setMessage(createMessage("Job marked as snagging.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not mark the job as snagging.", "error"));
    }
  }

  async function removeAdminJobSnagging(job) {
    if (!job?.id) return;
    try {
      await clearSnaggingJob(job.id);
      setAdminCompletePrompt(false);
      setMessage(createMessage("Snagging removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not remove snagging.", "error"));
    }
  }

  async function undoClientJobComplete(job) {
    if (!job?.id) return;
    try {
      await undoCompleteJob(job.id);
      setClientCompletePrompt(false);
      setMessage(createMessage("Job marked as not complete.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not undo completion.", "error"));
    }
  }

  async function undoAdminJobComplete(job) {
    if (!job?.id) return;
    try {
      await undoCompleteJob(job.id);
      setAdminCompletePrompt(false);
      setMessage(createMessage("Job marked as not complete.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not undo completion.", "error"));
    }
  }

  async function exportJob(job, setExportingState) {
    if (!job?.id) return;
    setExportingState(true);
    try {
      const printWindow = window.open("", "_blank");
      if (!printWindow) {
        throw new Error("Please allow pop-ups so the PDF export can open.");
      }

      const uploadedBy = [...new Set((job.photos || []).map((photo) => String(photo.uploadedByName || "").trim()).filter(Boolean))].join(", ") || "-";
      const installers = getInstallerDisplayList(job).join(", ") || "-";
      const firstPagePhotos = (job.photos || []).slice(0, 2);
      const remainingPhotoPages = [];
      for (let index = 2; index < (job.photos || []).length; index += 6) {
        remainingPhotoPages.push((job.photos || []).slice(index, index + 6));
      }

      const renderPhotoTile = (photo, index, extraClass = "") => `
        <figure class="photo-tile ${extraClass}">
          <div class="photo-frame">
            <img src="${escapeHtml(photo.url || buildJobPhotoUrl(job.id, photo.id))}" alt="Job photo ${index + 1}" />
          </div>
        </figure>
      `;

      const html = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(job.orderReference || job.customerName || "Job Export")}</title>
    <style>
      @font-face {
        font-family: 'Faricy';
        src: url('${window.location.origin}/fonts/FARICYNEW-MEDIUM.TTF') format('truetype');
        font-weight: 500;
      }
      @font-face {
        font-family: 'Faricy';
        src: url('${window.location.origin}/fonts/FARICYNEW-BOLD.TTF') format('truetype');
        font-weight: 700;
      }
      :root {
        color-scheme: light;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: 'Faricy', Arial, sans-serif;
        color: #1f2937;
        background: white;
      }
      .page {
        width: 210mm;
        min-height: 297mm;
        padding: 14mm;
        page-break-after: always;
      }
      .page:last-child { page-break-after: auto; }
      .header {
        display: flex;
        justify-content: space-between;
        align-items: start;
        gap: 16mm;
        margin-bottom: 8mm;
      }
      .header-copy h1 {
        margin: 0 0 2mm;
        font-size: 22px;
        line-height: 1.05;
      }
      .header-copy p {
        margin: 0;
        font-size: 12px;
      }
      .brand {
        width: 52mm;
        flex: 0 0 auto;
      }
      .brand img {
        display: block;
        width: 100%;
        height: auto;
      }
      .summary-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 4mm 8mm;
        margin-bottom: 8mm;
      }
      .summary-item {
        border: 1px solid #e2e8f0;
        border-radius: 4mm;
        padding: 3.2mm 3.6mm;
        min-height: 16mm;
      }
      .summary-item.wide {
        grid-column: 1 / -1;
      }
      .summary-item strong {
        display: block;
        margin-bottom: 1.2mm;
        color: #475569;
        font-size: 9px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
      }
      .summary-item span {
        font-size: 12px;
        line-height: 1.35;
      }
      .photo-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 6mm;
      }
      .photo-grid.first-page {
        margin-top: 4mm;
      }
      .photo-grid.extra-page {
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 5mm;
      }
      .photo-tile {
        margin: 0;
      }
      .photo-tile.first-page-photo .photo-frame {
        aspect-ratio: 1 / 1;
      }
      .photo-frame {
        border: 1px solid #dbe2ea;
        border-radius: 4mm;
        overflow: hidden;
        aspect-ratio: 1 / 1;
        background: #f8fafc;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .photo-frame img {
        width: 100%;
        height: 100%;
        object-fit: cover;
        display: block;
      }
      @page {
        size: A4;
        margin: 0;
      }
      @media print {
        body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      }
    </style>
  </head>
  <body>
    <section class="page">
      <div class="header">
        <div class="header-copy">
          <h1>${escapeHtml(job.customerName || "Job Export")}</h1>
          <p>${escapeHtml(job.description || "")}</p>
        </div>
        <div class="brand">
          <img src="${window.location.origin}/branding/signs-express-logo.svg" alt="Signs Express logo" />
        </div>
      </div>
      <div class="summary-grid">
        <div class="summary-item"><strong>Order Ref</strong><span>${escapeHtml(job.orderReference || "-")}</span></div>
        <div class="summary-item"><strong>Completion Date</strong><span>${escapeHtml(formatJobDate(job.date) || "-")}</span></div>
        <div class="summary-item"><strong>Job Type</strong><span>${escapeHtml(getJobTypeLabel(job))}</span></div>
        <div class="summary-item"><strong>Installers</strong><span>${escapeHtml(installers)}</span></div>
        <div class="summary-item"><strong>Contact</strong><span>${escapeHtml(job.contact || "-")}</span></div>
        <div class="summary-item"><strong>Number</strong><span>${escapeHtml(job.number || "-")}</span></div>
        <div class="summary-item wide"><strong>Address</strong><span>${escapeHtml(job.address || "-")}</span></div>
        <div class="summary-item wide"><strong>Photos Uploaded By</strong><span>${escapeHtml(uploadedBy)}</span></div>
      </div>
      ${firstPagePhotos.length ? `<div class="photo-grid first-page">${firstPagePhotos.map((photo, index) => renderPhotoTile(photo, index, "first-page-photo")).join("")}</div>` : ""}
    </section>
    ${remainingPhotoPages.map((pagePhotos) => `
      <section class="page">
        <div class="photo-grid extra-page">${pagePhotos.map((photo, index) => renderPhotoTile(photo, index)).join("")}</div>
      </section>
    `).join("")}
    <script>
      const images = Array.from(document.images);
      Promise.all(images.map((image) => image.complete ? Promise.resolve() : new Promise((resolve) => {
        image.addEventListener('load', resolve, { once: true });
        image.addEventListener('error', resolve, { once: true });
      }))).then(() => {
        setTimeout(() => {
          window.focus();
          window.print();
          setTimeout(() => {
            window.close();
          }, 300);
        }, 250);
      });
    </script>
  </body>
</html>`;

      printWindow.document.open();
      printWindow.document.write(html);
      printWindow.document.close();
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not export the job.", "error"));
    } finally {
      setExportingState(false);
    }
  }

  async function deleteJobPhoto(job, photoId) {
    if (!job?.id || !photoId) return;
    try {
      const response = await fetch(
        `/api/jobs/${encodeURIComponent(job.id)}/photos/${encodeURIComponent(photoId)}`,
        { method: "DELETE" }
      );
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not delete the photo.");
      }
      applyBoardPayloadToState(payload, job.id);
      setMessage(createMessage("Photo removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not delete the photo.", "error"));
    }
  }

  async function handleSubmit(event) {
    event.preventDefault();
    if (isClientMode) return;

    if (!form.customerName.trim()) {
      setMessage(createMessage("Enter the customer name.", "error"));
      return;
    }

    setSaving(true);

    try {
      const existingJob = editingId ? jobsById.get(editingId) : null;
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...existingJob,
          id: editingId || form.id || undefined,
          date: form.date,
          orderReference: form.orderReference.trim(),
          customerName: form.customerName.trim(),
          description: form.description.trim(),
          contact: form.contact.trim(),
          number: form.number.trim(),
          address: form.address.trim(),
          installers: form.installers,
          customInstaller: form.customInstaller.trim(),
          jobType: form.jobType,
          customJobType: form.customJobType.trim(),
          isPlaceholder: Boolean(form.isPlaceholder),
          notes: form.notes.trim()
        })
      });

      if (!response.ok) throw new Error("Could not save the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(createMessage(editingId ? "Job updated." : "Job added.", "success"));
      resetForm(form.date);
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save the job.", "error"));
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(jobId) {
    if (isClientMode) return;
    try {
      const response = await fetch(`/api/jobs/${jobId}`, { method: "DELETE" });
      if (!response.ok) throw new Error("Could not delete the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      if (editingId === jobId) resetForm();
      setMessage(createMessage("Job deleted.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete the job.", "error"));
    }
  }

  async function moveJobToDate(jobId, nextDate) {
    if (isClientMode) return;
    const job = jobsById.get(jobId);
    const normalizedNextDate = nextDate === UNSCHEDULED_DROP_ZONE ? "" : nextDate;
    if (!job || normalizedNextDate === undefined || normalizedNextDate === null || job.date === normalizedNextDate) {
      setDropDate("");
      setDraggingJobId("");
      return;
    }

    try {
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...job,
          date: normalizedNextDate
        })
      });

      if (!response.ok) throw new Error("Could not move the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      if (editingId === jobId) {
        setForm((current) => ({ ...current, date: normalizedNextDate }));
      }
      setMessage(
        createMessage(
          normalizedNextDate ? `Job moved to ${normalizedNextDate}.` : "Job moved to unscheduled.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not move the job.", "error"));
    } finally {
      setDropDate("");
      setDraggingJobId("");
    }
  }

  async function duplicateJobToDate(jobId, nextDate) {
    if (isClientMode) return;
    const job = jobsById.get(jobId);
    const normalizedNextDate = nextDate === UNSCHEDULED_DROP_ZONE ? "" : nextDate;
    if (!job || normalizedNextDate === undefined || normalizedNextDate === null) {
      setDuplicatingJobId("");
      setDropDate("");
      return;
    }

    try {
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...job,
          id: undefined,
          date: normalizedNextDate,
          isCompleted: false,
          completedAt: "",
          completedByUserId: "",
          completedByName: "",
          photos: []
        })
      });

      if (!response.ok) throw new Error("Could not duplicate the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(
        createMessage(
          normalizedNextDate ? `Job copied to ${normalizedNextDate}.` : "Job copied to unscheduled.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not duplicate the job.", "error"));
    } finally {
      setDuplicatingJobId("");
      setDropDate("");
      clearDragPreview();
    }
  }

  function getJobTypeMeta(jobType) {
    return JOB_TYPES.find((option) => option.value === jobType) || JOB_TYPES[JOB_TYPES.length - 1];
  }

  function getJobTypeLabel(job) {
    return job.jobType === "Other" ? job.customJobType || "Other" : job.jobType;
  }

  function toggleInstaller(value) {
    setForm((current) => ({
      ...current,
      installers: current.installers.includes(value)
        ? current.installers.filter((item) => item !== value)
        : [...current.installers, value]
    }));
  }

  function getInstallerMeta(value) {
    return INSTALLER_OPTIONS.find((option) => option.value === value) || { value, colorClass: "installer-custom" };
  }

  function getInstallerDisplayList(item) {
    const source = Array.isArray(item.installers)
      ? item.installers
      : typeof item.installers === "string" && item.installers.trim()
        ? item.installers.split(/[,/]+/).map((entry) => entry.trim()).filter(Boolean)
        : [];

    const visible = source.filter((entry) => entry !== "Custom");
    if (source.includes("Custom") && item.customInstaller) {
      visible.push(item.customInstaller);
    }
    return visible;
  }

  function getHolidayPersonColor(person) {
    return HOLIDAY_PERSON_COLORS[person] || "holiday-person-black";
  }

  function buildDragPreview(element) {
    const preview = element.cloneNode(true);
    preview.classList.add("drag-preview");
    preview.style.width = `${element.offsetWidth}px`;
    preview.style.position = "fixed";
    preview.style.top = "0";
    preview.style.left = "0";
    document.body.appendChild(preview);
    return preview;
  }

  function getTransparentDragImage() {
    if (!transparentDragImageRef.current) {
      const image = new Image();
      image.src =
        "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
      transparentDragImageRef.current = image;
    }

    return transparentDragImageRef.current;
  }

  function clearDragPreview() {
    if (dragPreviewRef.current) {
      dragPreviewRef.current.remove();
      dragPreviewRef.current = null;
    }
  }

  async function saveHoliday(date, person, duration = "Full Day", id) {
    if (isClientMode) return;
    try {
      const response = await fetch("/api/holidays", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id,
          date,
          person,
          duration
        })
      });
      if (!response.ok) throw new Error("Could not save holiday.");

        const payload = await response.json();
        setBoard(payload.board);
        setJobs(payload.jobs);
        setHolidays(payload.holidays);
        if (showHolidays) {
          await refreshHolidayData();
        }
        setMessage(createMessage("Holiday updated.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save holiday.", "error"));
    }
  }

  async function deleteHoliday(holidayId, options = {}) {
    if (isClientMode) return;
    try {
      const params = new URLSearchParams();
      if (options.date) params.set("date", options.date);
      if (options.person) params.set("person", options.person);
      const url = params.toString() ? `/api/holidays/${holidayId}?${params.toString()}` : `/api/holidays/${holidayId}`;
      const response = await fetch(url, { method: "DELETE" });
      if (!response.ok) throw new Error("Could not delete holiday.");

        const payload = await response.json();
        setBoard(payload.board);
        setJobs(payload.jobs);
        setHolidays(payload.holidays);
        if (showHolidays) {
          await refreshHolidayData();
        }
        setMessage(createMessage("Holiday removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete holiday.", "error"));
    }
  }

  async function cycleHoliday(date, person) {
    if (isClientMode) return;
    const normalizedPerson = getHolidayStaffIdentityKey(person);
    const matchingEntries = holidays.filter(
      (item) =>
        String(item.date || "") === String(date || "") &&
        getHolidayStaffIdentityKey(item.person) === normalizedPerson
    );
    const existing =
      matchingEntries.find((item) => !isBirthdayHoliday(item)) ||
      matchingEntries[0] ||
      null;

    if (existing && isBirthdayHoliday(existing)) return;

    if (!existing) {
      await saveHoliday(date, person, "Full Day");
      return;
    }

    if (existing.duration === "Full Day") {
      await saveHoliday(date, person, "Morning", existing.id);
      return;
    }

    if (existing.duration === "Morning") {
      await saveHoliday(date, person, "Afternoon", existing.id);
      return;
    }

    await deleteHoliday(existing.id, { date, person });
  }

  async function submitHolidayEvent() {
    if (!canEditHolidays(currentUser)) return;
    setHolidayEventSaving(true);

    try {
      const response = await fetch("/api/holiday-events", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(holidayEventForm)
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not save calendar event.");
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayEventForm(EMPTY_HOLIDAY_EVENT_FORM);
      setHolidayEventOpen(false);
      setMessage(createMessage("Calendar event updated.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not save calendar event.", "error"));
    } finally {
      setHolidayEventSaving(false);
    }
  }

  async function deleteHolidayEvent(eventId) {
    if (!canEditHolidays(currentUser)) return;
    setHolidayEventSaving(true);

    try {
      const response = await fetch(
        `/api/holiday-events/${encodeURIComponent(eventId)}?yearStart=${encodeURIComponent(holidayYearStart)}`,
        { method: "DELETE" }
      );
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not delete calendar event.");
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayEventForm(EMPTY_HOLIDAY_EVENT_FORM);
      setHolidayEventOpen(false);
      setMessage(createMessage("Calendar event removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not delete calendar event.", "error"));
    } finally {
      setHolidayEventSaving(false);
    }
  }

  async function handleHolidayChipClick(date, person, closePicker = false) {
    await cycleHoliday(date, person);
    if (closePicker) {
      setActiveHolidayDate("");
    }
  }

  async function handleManualRefresh() {
    try {
      const [boardResponse] = await Promise.all([fetch("/api/board"), refreshData()]);
      if (!boardResponse.ok) throw new Error("Could not refresh the board.");
      setBoard(await boardResponse.json());
      setMessage(createMessage("Board refreshed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not refresh the board.", "error"));
    }
  }

  if (!authChecked) {
    return (
      <div className="app-shell">
        <div className="page">
          <section className="panel auth-panel">
            <div className="board-loading">Checking login...</div>
          </section>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className="app-shell">
        <div className="page">
          <section className="panel auth-panel">
            <div className="auth-brand">
              <img className="auth-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
            </div>
            <form className="job-form auth-form" onSubmit={handleLogin}>
              <label>
                Username
                <input
                  type="text"
                  value={loginDisplayName}
                  placeholder="Enter your full name"
                  onChange={(event) => setLoginDisplayName(event.target.value)}
                />
              </label>

              <label>
                Password
                <input
                  type="password"
                  value={loginPassword}
                  placeholder="Enter your password"
                  onChange={(event) => setLoginPassword(event.target.value)}
                />
              </label>

              {loginError ? <div className="flash error">{loginError}</div> : null}

              <div className="form-actions">
                <button className="primary-button" type="submit" disabled={loginLoading || !loginDisplayName || !loginPassword}>
                  {loginLoading ? "Signing in..." : "Sign in"}
                </button>
              </div>
            </form>
          </section>
        </div>
      </div>
    );
  }

  if (showHostLanding) {
    return (
        <HostLandingPage
          currentUser={currentUser}
          onLogout={handleLogout}
          users={loginUsers}
          savingKey={permissionSavingKey}
          onChangePermission={handlePermissionChange}
          onUpdateAttendanceProfile={handleUpdateAttendanceProfile}
          onCreateUser={handleCreateUser}
          onResetPassword={handleResetUserPassword}
          onDeleteUser={handleDeleteUser}
          notifications={notifications}
        />
    );
  }

  if (showClientLanding) {
    return (
      <ClientLandingPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
      />
    );
  }

  if (showNotifications) {
    return (
      <NotificationsPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        onOpenNotification={openNotification}
        onMarkNotificationRead={markNotificationRead}
        onMarkAllNotificationsRead={markAllNotificationsRead}
      />
    );
  }

  if (showMileage) {
    return (
      <MileagePage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        onRefreshNotifications={refreshNotifications}
      />
    );
  }

  if (showVanEstimator) {
    return (
      <VinylEstimatorPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
      />
    );
  }

  if (showRamsLogic) {
    return (
      <RamsLogicPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
      />
    );
  }

  if (showRams) {
    return (
      <RamsPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
      />
    );
  }

  if (showAttendance) {
    return (
      <AttendancePage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        attendanceData={attendanceData}
        loading={attendanceLoading}
        attendanceMonthId={attendanceMonthId}
        setAttendanceMonthId={setAttendanceMonthId}
        attendanceSavingKey={attendanceSavingKey}
        attendanceNoteSavingKey={attendanceNoteSavingKey}
        attendanceFocusDate={attendanceNotificationDate}
        onSaveAttendanceEntry={saveAttendanceEntry}
        onSubmitAttendanceExplanation={submitAttendanceExplanation}
      />
    );
  }

  if (showHolidays) {
    return (
      <HolidaysPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        holidays={holidays}
        holidayRequests={holidayRequests}
        approvedHolidayRequests={approvedHolidayRequests}
        holidayStaff={holidayStaff}
        holidayAllowances={holidayAllowances}
        holidayEvents={holidayEvents}
        holidayRows={holidayRows}
        holidayYearStart={holidayYearStart}
        holidayYearLabel={holidayYearLabel}
        currentHolidayYearStart={currentHolidayYearStart}
        setHolidayYearStart={setHolidayYearStart}
          holidayRequestOpen={holidayRequestOpen}
          setHolidayRequestOpen={setHolidayRequestOpen}
          holidayRequestForm={holidayRequestForm}
          setHolidayRequestForm={setHolidayRequestForm}
          holidayRequestSaving={holidayRequestSaving}
          holidayCancelOpen={holidayCancelOpen}
          setHolidayCancelOpen={setHolidayCancelOpen}
          holidayCancelForm={holidayCancelForm}
          setHolidayCancelForm={setHolidayCancelForm}
          holidayEventOpen={holidayEventOpen}
          setHolidayEventOpen={setHolidayEventOpen}
          holidayEventForm={holidayEventForm}
        setHolidayEventForm={setHolidayEventForm}
        holidayEventSaving={holidayEventSaving}
        holidayAllowanceSavingKey={holidayAllowanceSavingKey}
        onChangeHolidayAllowanceDraft={changeHolidayAllowanceDraft}
        onSaveHolidayAllowance={saveHolidayAllowance}
        onToggleHolidayDate={cycleHoliday}
        onSubmitHolidayEvent={submitHolidayEvent}
          onDeleteHolidayEvent={deleteHolidayEvent}
          onSubmitHolidayRequest={submitHolidayRequest}
          onReviewHolidayRequest={reviewHolidayRequest}
          onCancelHolidayRequest={cancelHolidayRequest}
        />
      );
    }

  if (showInstallerDirectory) {
    return <InstallerDirectoryHost currentUser={currentUser} onLogout={handleLogout} readOnly={!installerEditable} />;
  }

  return (
    <div className={`app-shell ${isClientMode ? "client-mode" : "editor-mode"}`}>
      <div className="page">
        <MainNavBar
          currentUser={currentUser}
          active="board"
          onLogout={handleLogout}
          notifications={notifications}
        />

        <div className="layout">
          <section className="panel board-panel board-panel-full">
            {loading || !board ? (
              <div className="board-loading">Loading the shared installation board...</div>
            ) : (
              <div className="board board-with-history">
                <div className="board-history-launch board-history-launch-top">
                  <div className="board-history-actions">
                    <button
                      className="ghost-button board-history-button"
                      type="button"
                      onClick={() => setPreviousMonthDepth((current) => Math.min(6, current + 1))}
                    >
                      Previous months
                    </button>
                    <button className="ghost-button board-history-button" type="button" onClick={resetBoardWindow}>
                      Current month
                    </button>
                  </div>
                </div>

                {board.weeks.map((week) => (
                  <section key={week.id} className="week-block">
                    <header className="week-header">
                      <strong>{week.label}</strong>
                    </header>

                    {week.rows.map((row) => (
                      <article
                        key={row.isoDate}
                        className={[
                          "board-row",
                          row.isToday ? "is-today" : "",
                          row.bankHoliday ? "is-bank-holiday" : "",
                          row.isPast ? "is-past" : "",
                          dropDate === row.isoDate ? "is-drop-target" : ""
                        ].join(" ").trim()}
                        onDragOver={(event) => {
                          if (isClientMode) return;
                          event.preventDefault();
                          if (draggingJobId) setDropDate(row.isoDate);
                        }}
                        onDragLeave={() => {
                          if (dropDate === row.isoDate) setDropDate("");
                        }}
                        onDrop={(event) => {
                          if (isClientMode) return;
                          event.preventDefault();
                          const duplicateJobId = event.dataTransfer.getData("job-copy");
                          if (duplicateJobId || duplicatingJobId) {
                            duplicateJobToDate(duplicateJobId || duplicatingJobId, row.isoDate);
                            return;
                          }
                          const jobId = event.dataTransfer.getData("text/plain") || draggingJobId;
                          moveJobToDate(jobId, row.isoDate);
                        }}
                      >
                          <div
                            className="date-cell"
                          onClick={() => {
                            if (!isClientMode) {
                              setActiveHolidayDate((current) => (current === row.isoDate ? "" : row.isoDate));
                            }
                          }}
                          title={row.fullDateLabel}
                          >
                            <div className="date-heading">
                              <span className="date-day">{row.dayLabel}</span>
                              <strong className="date-number">{row.dayNumber}</strong>
                            </div>
                            {row.isToday ? <span className="date-today-pill">Today</span> : null}
                            {isClientMode && row.staffHolidays.length ? (
                            <div className="mobile-holiday-inline">
                              {row.staffHolidays.map((holiday) => (
                                <span
                                  key={`mobile-${holiday.id}`}
                                  className={`mobile-holiday-chip ${getHolidayPersonColor(holiday.person)} ${isBirthdayHoliday(holiday) ? "holiday-birthday-token" : ""}`}
                                >
                                  {getHolidayDisplayToken(holiday.person)}
                                  {holiday.duration === "Morning" ? " AM" : holiday.duration === "Afternoon" ? " PM" : ""}
                                </span>
                              ))}
                            </div>
                          ) : null}
                          {!isClientMode && row.staffHolidays.length ? (
                            <div className="date-holiday-summary" onClick={(event) => event.stopPropagation()}>
                              {row.staffHolidays.map((holiday) => {
                                const durationLabel =
                                  holiday.duration === "Morning"
                                    ? ".AM"
                                    : holiday.duration === "Afternoon"
                                      ? ".PM"
                                      : "";
                                return (
                                  <button
                                    key={`summary-${holiday.id}`}
                                    type="button"
                                    className={`date-holiday-chip ${getHolidayPersonColor(holiday.person)} ${isBirthdayHoliday(holiday) ? "holiday-birthday-token" : ""} active`}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      if (!isBirthdayHoliday(holiday)) {
                                        handleHolidayChipClick(row.isoDate, holiday.person, false);
                                      }
                                    }}
                                    disabled={isBirthdayHoliday(holiday)}
                                  >
                                    {getHolidayDisplayToken(holiday.person)}{durationLabel}
                                  </button>
                                );
                              })}
                            </div>
                          ) : null}
                          {!isClientMode && activeHolidayDate === row.isoDate ? (
                              <div className="date-holiday-popover" onClick={(event) => event.stopPropagation()}>
                                {holidayStaff.map((entry) => {
                                  const name = entry.person;
                                  const existing = row.staffHolidays.find((holiday) => holiday.person === name);
                                  const durationLabel =
                                    existing?.duration === "Morning"
                                      ? ".AM"
                                      : existing?.duration === "Afternoon"
                                      ? ".PM"
                                      : "";
                                const initials = getHolidayDisplayToken(name);
                                return (
                                  <button
                                    key={`${row.isoDate}-${name}`}
                                    type="button"
                                    className={`date-holiday-chip ${getHolidayPersonColor(name)} ${existing ? "active" : ""} ${isBirthdayHoliday(existing) ? "holiday-birthday-token" : ""}`}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      if (!isBirthdayHoliday(existing)) {
                                        handleHolidayChipClick(row.isoDate, name, true);
                                      }
                                    }}
                                    disabled={isBirthdayHoliday(existing)}
                                  >
                                    {initials}{durationLabel}
                                  </button>
                                );
                              })}
                            </div>
                          ) : null}
                          {!isClientMode && row.bankHoliday ? <span className="date-holiday-chip date-bank-holiday">{row.bankHoliday}</span> : null}
                          {Array.isArray(row.holidayEvents) && row.holidayEvents.length ? (
                            <div className="date-calendar-events">
                              {row.holidayEvents.map((event) => (
                                <span key={`board-event-${event.id}`} className="date-calendar-event-chip">
                                  {event.title}
                                </span>
                              ))}
                            </div>
                          ) : null}
                        </div>

                        <div className="jobs-cell">
                          <button
                            type="button"
                            className={`jobs-lane-button ${row.jobs.length === 0 ? "is-empty" : ""}`}
                            disabled={isClientMode}
                            onClick={() => {
                              if (isClientMode) return;
                              openJobModal(row.isoDate, {
                                ...EMPTY_FORM,
                                date: row.isoDate,
                                jobType: form.jobType || "Install"
                              });
                            }}
                          >
                            {row.jobs.length === 0 ? <span className="muted">No jobs booked</span> : <span className="lane-add-label">{isClientMode ? "View only" : "Click anywhere here to add another job"}</span>}
                          </button>

                          {row.jobs.length > 0 ? (
                            <div className="job-stack">
                                {row.jobs.map((job) =>
                                  renderJobCardContent({
                                    job,
                                    isCondensed: row.isPast,
                                    isClientMode,
                                    draggingJobId,
                                  getJobTypeMeta,
                                  getJobTypeLabel,
                                  getInstallerDisplayList,
                                  getInstallerMeta,
                                  editJob,
                                  handleDelete,
                                  setActiveClientJob,
                                  buildDragPreview,
                                  getTransparentDragImage,
                                  clearDragPreview,
                                  dragPreviewRef,
                                  dragPositionRef,
                                  setDraggingJobId,
                                  duplicatingJobId,
                                  setDuplicatingJobId,
                                  setDropDate
                                })
                              )}
                            </div>
                          ) : null}
                        </div>
                      </article>
                    ))}
                  </section>
                ))}

                <div className="board-history-launch board-history-launch-bottom">
                  <div className="board-history-actions">
                    <button
                      className="ghost-button board-history-button"
                      type="button"
                      onClick={() => setFutureMonthDepth((current) => Math.min(6, current + 1))}
                    >
                      Future months
                    </button>
                  </div>
                </div>

                <section
                  className={`unscheduled-section panel ${dropDate === UNSCHEDULED_DROP_ZONE ? "is-drop-target" : ""}`}
                  onDragOver={(event) => {
                    if (isClientMode) return;
                    event.preventDefault();
                    if (draggingJobId || duplicatingJobId) setDropDate(UNSCHEDULED_DROP_ZONE);
                  }}
                  onDragLeave={() => {
                    if (dropDate === UNSCHEDULED_DROP_ZONE) setDropDate("");
                  }}
                  onDrop={(event) => {
                    if (isClientMode) return;
                    event.preventDefault();
                    const duplicateJobId = event.dataTransfer.getData("job-copy");
                    if (duplicateJobId || duplicatingJobId) {
                      duplicateJobToDate(duplicateJobId || duplicatingJobId, UNSCHEDULED_DROP_ZONE);
                      return;
                    }
                    const jobId = event.dataTransfer.getData("text/plain") || draggingJobId;
                    moveJobToDate(jobId, UNSCHEDULED_DROP_ZONE);
                  }}
                >
                  <div className="unscheduled-head">
                    <div>
                      <h3>Unscheduled</h3>
                      <p>Keep jobs here until you have a firm date, then drag them into the calendar or edit the date.</p>
                    </div>
                    {!isClientMode ? (
                      <button
                        className="ghost-button"
                        type="button"
                        onClick={() => {
                          openJobModal("Unscheduled", {
                            ...EMPTY_FORM,
                            date: "",
                            jobType: form.jobType || "Install"
                          });
                        }}
                      >
                        Add unscheduled job
                      </button>
                    ) : null}
                  </div>

                  {board.unscheduled?.length ? (
                    <div className="job-stack unscheduled-job-stack">
                        {board.unscheduled.map((job) =>
                          renderJobCardContent({
                            job,
                            isCondensed: false,
                            isClientMode,
                          draggingJobId,
                          getJobTypeMeta,
                          getJobTypeLabel,
                          getInstallerDisplayList,
                          getInstallerMeta,
                          editJob,
                          handleDelete,
                          setActiveClientJob,
                          buildDragPreview,
                          getTransparentDragImage,
                          clearDragPreview,
                          dragPreviewRef,
                          dragPositionRef,
                          setDraggingJobId,
                          duplicatingJobId,
                          setDuplicatingJobId,
                          setDropDate
                        })
                      )}
                    </div>
                  ) : (
                    <div className="unscheduled-empty">No unscheduled jobs.</div>
                  )}
                </section>
              </div>
            )}
          </section>
        </div>
      </div>
      {!isClientMode && jobModalDate ? (
        <div
          className="modal-backdrop"
          onPointerDown={handleBackdropPointerDown}
          onClick={(event) => handleBackdropClick(event, () => resetForm())}
        >
          <div className="modal job-modal" onPointerDown={() => { backdropPointerStartedRef.current = false; }} onClick={(event) => event.stopPropagation()}>
            <div className="modal-head job-modal-head">
              <button className="icon-button" type="button" onClick={() => resetForm()}>
                x
              </button>
            </div>

            <form className="job-form job-form-scroll" onSubmit={handleSubmit}>
              <label>
                Date
                <input
                  type="date"
                  value={form.date}
                  onChange={(event) => setForm((current) => ({ ...current, date: event.target.value }))}
                />
              </label>

              <div className="corebridge-lookup-bar">
                <button className="ghost-button" type="button" onClick={() => openOrderLookup()}>
                  Find order
                </button>
                <span className="muted">Pull live order details from CoreBridge where available.</span>
              </div>

              <label>
                Order reference
                <div className="order-reference-row">
                  <input
                    type="text"
                    value={form.orderReference}
                    onChange={(event) => setForm((current) => ({ ...current, orderReference: event.target.value }))}
                  />
                  <button
                    type="button"
                    className={`placeholder-toggle-button ${form.isPlaceholder ? "active" : ""}`}
                    onClick={() => setForm((current) => ({ ...current, isPlaceholder: !current.isPlaceholder }))}
                  >
                    Add as Placeholder
                  </button>
                </div>
              </label>

              <label>
                Customer name
                <input
                  type="text"
                  value={form.customerName}
                  onChange={(event) => setForm((current) => ({ ...current, customerName: event.target.value }))}
                />
              </label>

              <label>
                Description
                <input
                  type="text"
                  value={form.description}
                  onChange={(event) => setForm((current) => ({ ...current, description: event.target.value }))}
                />
              </label>

              <div className="split-fields">
                <label>
                  Contact
                  <input
                    type="text"
                    value={form.contact}
                    onChange={(event) => setForm((current) => ({ ...current, contact: event.target.value }))}
                  />
                </label>

                <label>
                  Number
                  <input
                    type="text"
                    value={form.number}
                    onChange={(event) => setForm((current) => ({ ...current, number: event.target.value }))}
                  />
                </label>
              </div>

              <label>
                Address
                <input
                  type="text"
                  value={form.address}
                  onChange={(event) => setForm((current) => ({ ...current, address: event.target.value }))}
                />
              </label>

              <label>
                Installers
                <div className="installer-picker">
                  {INSTALLER_OPTIONS.map((option) => (
                    <button
                      key={option.value}
                      type="button"
                      className={`installer-chip ${option.colorClass} ${form.installers.includes(option.value) ? "active" : ""}`}
                      onClick={() => toggleInstaller(option.value)}
                    >
                      {option.value}
                    </button>
                  ))}
                </div>
              </label>

              {form.installers.includes("Custom") ? (
                <label>
                  Custom installer
                  <input
                    type="text"
                    value={form.customInstaller}
                    onChange={(event) => setForm((current) => ({ ...current, customInstaller: event.target.value }))}
                  />
                </label>
              ) : null}

              <label>
                Job type
                <select
                  value={form.jobType}
                  onChange={(event) => setForm((current) => ({ ...current, jobType: event.target.value }))}
                >
                  {JOB_TYPES.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.value}
                    </option>
                  ))}
                </select>
              </label>

              {form.jobType === "Other" ? (
                <label>
                  Other job type
                  <input
                    type="text"
                    value={form.customJobType}
                    onChange={(event) => setForm((current) => ({ ...current, customJobType: event.target.value }))}
                  />
                </label>
              ) : null}

              <label>
                Notes
                <textarea
                  rows="4"
                  value={form.notes}
                  onChange={(event) => setForm((current) => ({ ...current, notes: event.target.value }))}
                />
              </label>

              {activeAdminJob ? (
                <>
                  <div className="client-job-summary admin-job-summary">
                    {activeAdminJob.orderReference ? <span className="job-summary-pill">{activeAdminJob.orderReference}</span> : null}
                    <span className={`job-summary-pill ${activeAdminJob.isCompleted ? "is-complete" : ""}`}>
                      {activeAdminJob.isCompleted ? "Completed" : getJobTypeLabel(activeAdminJob)}
                    </span>
                    {activeAdminJob.isSnagging ? <span className="job-summary-pill is-snagging">Snagging</span> : null}
                    {activeAdminJob.isPlaceholder ? <span className="job-summary-pill is-placeholder">Placeholder</span> : null}
                    {Array.isArray(activeAdminJob.photos) && activeAdminJob.photos.length ? (
                      <span className="job-summary-pill is-photos">{activeAdminJob.photos.length} photo{activeAdminJob.photos.length === 1 ? "" : "s"}</span>
                    ) : null}
                  </div>

                  <div className="detail-grid client-detail-grid admin-job-detail-grid">
                    {activeAdminJob.completedAt ? (
                      <div className="detail-card">
                        <strong>Completed At</strong>
                        <p>{formatDateTime(activeAdminJob.completedAt)}</p>
                      </div>
                    ) : null}
                    {activeAdminJob.completedByName ? (
                      <div className="detail-card">
                        <strong>Completed By</strong>
                        <p>{activeAdminJob.completedByName}</p>
                      </div>
                    ) : null}
                    <div className="detail-card detail-card-wide">
                      <strong>Photos</strong>
                      {Array.isArray(activeAdminJob.photos) && activeAdminJob.photos.length ? (
                        <div className="job-photo-grid">
                          {activeAdminJob.photos.map((photo) => (
                            <div key={photo.id} className="job-photo-tile">
                              <a
                                className="job-photo-link"
                                href={photo.url || buildJobPhotoUrl(activeAdminJob.id, photo.id)}
                                target="_blank"
                                rel="noreferrer"
                                onClick={(event) => event.stopPropagation()}
                              >
                                <img
                                  src={photo.url || buildJobPhotoUrl(activeAdminJob.id, photo.id)}
                                  alt={photo.fileName || "Job photo"}
                                  loading="lazy"
                                />
                              </a>
                              <div className="job-photo-meta">
                                <small>{photo.uploadedByName || "Uploaded photo"}</small>
                                <button
                                  className="text-button danger"
                                  type="button"
                                  onClick={() => {
                                    deleteJobPhoto(activeAdminJob, photo.id);
                                  }}
                                >
                                  Delete
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p>No photos uploaded yet.</p>
                      )}
                    </div>
                  </div>

                  </>
                ) : null}

              <div className="form-actions job-form-actions">
                {activeAdminJob && !activeAdminJob.isCompleted && !adminCompletePrompt ? (
                  <button
                    className="success-button"
                    type="button"
                    onClick={() => setAdminCompletePrompt(true)}
                    disabled={adminPhotoUploading || adminExporting}
                  >
                    Mark as Complete
                  </button>
                ) : null}
                {activeAdminJob && !activeAdminJob.isCompleted && !activeAdminJob.isSnagging && !adminCompletePrompt ? (
                  <button
                    className="snagging-button"
                    type="button"
                    onClick={() => markAdminJobSnagging(activeAdminJob)}
                    disabled={adminPhotoUploading || adminExporting}
                  >
                    Snagging
                  </button>
                ) : null}
                {activeAdminJob && activeAdminJob.isSnagging && !adminCompletePrompt ? (
                  <button
                    className="snagging-button is-active"
                    type="button"
                    onClick={() => removeAdminJobSnagging(activeAdminJob)}
                    disabled={adminPhotoUploading || adminExporting}
                  >
                    Remove Snagging
                  </button>
                ) : null}
                <button className="primary-button" type="submit" disabled={saving}>
                  {saving ? "Saving..." : editingId ? "Update job" : "Add job"}
                </button>
                <button className="ghost-button" type="button" onClick={() => resetForm()}>
                  Cancel
                </button>
                {activeAdminJob?.isCompleted ? (
                  <>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => adminPhotoInputRef.current?.click()}
                      disabled={adminPhotoUploading}
                    >
                      {adminPhotoUploading ? "Uploading..." : "Upload photos"}
                    </button>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => undoAdminJobComplete(activeAdminJob)}
                      disabled={adminPhotoUploading || adminExporting}
                    >
                      Undo complete
                    </button>
                  </>
                ) : null}
                {activeAdminJob ? (
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={() => exportJob(activeAdminJob, setAdminExporting)}
                    disabled={adminExporting || adminPhotoUploading}
                  >
                    {adminExporting ? "Exporting..." : "Export"}
                  </button>
                ) : null}
                {activeAdminJob && !activeAdminJob.isCompleted && adminCompletePrompt ? (
                  <div className="client-complete-prompt form-complete-prompt">
                    <span>Would you like to upload job photos?</span>
                    <div className="client-complete-prompt-actions">
                      <button
                        className="ghost-button"
                        type="button"
                        onClick={() => markAdminJobComplete(activeAdminJob)}
                        disabled={adminPhotoUploading}
                      >
                        No
                      </button>
                      <button
                        className="success-button"
                        type="button"
                        onClick={() => adminPhotoInputRef.current?.click()}
                        disabled={adminPhotoUploading}
                      >
                        {adminPhotoUploading ? "Uploading..." : "Yes"}
                      </button>
                    </div>
                  </div>
                ) : null}
              </div>
            </form>
            <input
              ref={adminPhotoInputRef}
              className="visually-hidden"
              type="file"
              accept="image/*"
              multiple
              onChange={async (event) => {
                const files = Array.from(event.target.files || []);
                if (!files.length || !activeAdminJob) return;
                await markAdminJobComplete(activeAdminJob, files);
              }}
            />
          </div>
        </div>
      ) : null}
      {!isClientMode && orderLookupOpen ? (
        <div
          className="modal-backdrop"
          onPointerDown={handleBackdropPointerDown}
          onClick={(event) => handleBackdropClick(event, () => setOrderLookupOpen(false))}
        >
          <div className="modal order-lookup-modal" onPointerDown={() => { backdropPointerStartedRef.current = false; }} onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>Find CoreBridge Order</h3>
                <p>Search live orders and copy the matching details into this job.</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setOrderLookupOpen(false)}>
                x
              </button>
            </div>

            <div className="order-lookup-toolbar">
              <input
                type="text"
                value={orderLookupQuery}
                placeholder="Search by order ref, customer, phone or address"
                onChange={(event) => setOrderLookupQuery(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    event.preventDefault();
                    searchCoreBridgeOrders(orderLookupQuery);
                  }
                }}
              />
              <button className="primary-button" type="button" onClick={() => searchCoreBridgeOrders(orderLookupQuery)} disabled={orderLookupLoading}>
                {orderLookupLoading ? "Searching..." : "Search"}
              </button>
            </div>

            {orderLookupError ? <div className="flash error">{orderLookupError}</div> : null}

            <div className="order-lookup-results">
              {orderLookupLoading ? (
                <div className="board-loading compact">Looking up CoreBridge orders...</div>
              ) : orderLookupResults.length ? (
                orderLookupResults.map((order) => (
                  <div
                    key={`${order.id}-${order.orderReference}-${order.customerName}`}
                    className="order-result-card"
                  >
                    <div className="order-result-top">
                      <strong>{order.orderReference || "No order ref"}</strong>
                      {order.status ? <span className="job-tag job-type-other">{order.status}</span> : null}
                    </div>
                    <p className="order-result-customer">{order.customerName || "Unnamed customer"}</p>
                    <p>{order.description || "No description"}</p>
                    <div className="order-result-meta">
                      <span><b>Contact:</b> {order.contact || "-"}</span>
                      <span><b>Number:</b> {order.number || "-"}</span>
                    </div>
                    <p className="order-result-address"><b>Address:</b> {order.address || "-"}</p>
                    <div className="order-result-actions">
                      <button className="primary-button" type="button" onClick={() => applyCoreBridgeOrder(order)}>
                        Use this order
                      </button>
                    </div>
                  </div>
                ))
              ) : (
                <div className="board-loading compact">No CoreBridge orders found yet.</div>
              )}
            </div>
          </div>
        </div>
      ) : null}
      {isClientMode && activeClientJob ? (
        <div
          className="modal-backdrop"
          onPointerDown={handleBackdropPointerDown}
          onClick={(event) => handleBackdropClick(event, () => setActiveClientJob(null))}
        >
          <div className="modal client-detail-modal" onPointerDown={() => { backdropPointerStartedRef.current = false; }} onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>{activeClientJob.customerName}</h3>
                <p>{activeClientJob.description || "No description"}</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setActiveClientJob(null)}>
                x
              </button>
            </div>
            <div className="client-detail-scroll">
              <div className="client-job-summary">
                {activeClientJob.orderReference ? <span className="job-summary-pill">{activeClientJob.orderReference}</span> : null}
                <span className={`job-summary-pill ${activeClientJob.isCompleted ? "is-complete" : ""}`}>
                  {activeClientJob.isCompleted ? "Completed" : getJobTypeLabel(activeClientJob)}
                </span>
                {activeClientJob.isSnagging ? <span className="job-summary-pill is-snagging">Snagging</span> : null}
                {activeClientJob.isPlaceholder ? <span className="job-summary-pill is-placeholder">Placeholder</span> : null}
                {Array.isArray(activeClientJob.photos) && activeClientJob.photos.length ? (
                  <span className="job-summary-pill is-photos">{activeClientJob.photos.length} photo{activeClientJob.photos.length === 1 ? "" : "s"}</span>
                ) : null}
              </div>
              <div className="detail-grid client-detail-grid">
                <div className="detail-card">
                  <strong>Contact</strong>
                  <p>{activeClientJob.contact || "-"}</p>
                </div>
                <div className="detail-card">
                  <strong>Number</strong>
                  <p>{activeClientJob.number || "-"}</p>
                </div>
                <div className="detail-card">
                  <strong>Installers</strong>
                  <p>{getInstallerDisplayList(activeClientJob).join(", ") || "-"}</p>
                </div>
                <div className="detail-card">
                  <strong>Placeholder</strong>
                  <p>{activeClientJob.isPlaceholder ? "Yes" : "No"}</p>
                </div>
                {activeClientJob.completedAt ? (
                  <div className="detail-card">
                    <strong>Completed At</strong>
                    <p>{formatDateTime(activeClientJob.completedAt)}</p>
                  </div>
                ) : null}
                {activeClientJob.completedByName ? (
                  <div className="detail-card">
                    <strong>Completed By</strong>
                    <p>{activeClientJob.completedByName}</p>
                  </div>
                ) : null}
                <div className="detail-card detail-card-wide">
                  <strong>Address</strong>
                  <p>{activeClientJob.address || "-"}</p>
                </div>
                <div className="detail-card detail-card-wide">
                  <strong>Notes</strong>
                  <p>{activeClientJob.notes || "-"}</p>
                </div>
                <div className="detail-card detail-card-wide">
                  <strong>Photos</strong>
                  {Array.isArray(activeClientJob.photos) && activeClientJob.photos.length ? (
                    <div className="job-photo-grid">
                      {activeClientJob.photos.map((photo) => (
                        <div key={photo.id} className="job-photo-tile">
                          <a
                            className="job-photo-link"
                            href={photo.url || buildJobPhotoUrl(activeClientJob.id, photo.id)}
                            target="_blank"
                            rel="noreferrer"
                            onClick={(event) => event.stopPropagation()}
                          >
                            <img
                              src={photo.url || buildJobPhotoUrl(activeClientJob.id, photo.id)}
                              alt={photo.fileName || "Job photo"}
                              loading="lazy"
                            />
                          </a>
                          <div className="job-photo-meta">
                            <small>{photo.uploadedByName || "Uploaded photo"}</small>
                            <button
                              className="text-button danger"
                              type="button"
                              onClick={(event) => {
                                event.stopPropagation();
                                deleteJobPhoto(activeClientJob, photo.id);
                              }}
                            >
                              Delete
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p>No photos uploaded yet.</p>
                  )}
                </div>
              </div>
              <div className="client-job-actions">
                {!activeClientJob.isCompleted ? (
                  <>
                    {!clientCompletePrompt ? (
                      <button
                        className="primary-button"
                        type="button"
                        onClick={() => setClientCompletePrompt(true)}
                        disabled={clientPhotoUploading || clientExporting}
                      >
                        Mark as Complete
                      </button>
                    ) : (
                      <div className="client-complete-prompt">
                        <span>Would you like to upload job photos?</span>
                        <div className="client-complete-prompt-actions">
                          <button
                            className="ghost-button"
                            type="button"
                            onClick={() => markClientJobComplete(activeClientJob)}
                            disabled={clientPhotoUploading}
                          >
                            No
                          </button>
                          <button
                            className="primary-button"
                            type="button"
                            onClick={() => clientPhotoInputRef.current?.click()}
                            disabled={clientPhotoUploading}
                          >
                            {clientPhotoUploading ? "Uploading..." : "Yes"}
                          </button>
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => clientPhotoInputRef.current?.click()}
                      disabled={clientPhotoUploading}
                    >
                      {clientPhotoUploading ? "Uploading..." : "Upload photos"}
                    </button>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => undoClientJobComplete(activeClientJob)}
                      disabled={clientPhotoUploading || clientExporting}
                    >
                      Undo complete
                    </button>
                  </>
                )}

                <button
                  className="ghost-button"
                  type="button"
                  onClick={() => exportJob(activeClientJob, setClientExporting)}
                  disabled={clientExporting || clientPhotoUploading}
                >
                  {clientExporting ? "Exporting..." : "Export"}
                </button>
              </div>
            </div>
            <input
              ref={clientPhotoInputRef}
              className="visually-hidden"
              type="file"
              accept="image/*"
              multiple
              onChange={async (event) => {
                const files = Array.from(event.target.files || []);
                if (!files.length) return;
                await markClientJobComplete(activeClientJob, files);
              }}
            />
          </div>
        </div>
      ) : null}
    </div>
  );
}
