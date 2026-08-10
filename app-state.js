// ============================================================
// MICROSOFT AUTHENTICATION CONFIGURATION
// ============================================================
const MSAL_CONFIG = {
  auth: {
    clientId: "8d523e65-0163-49c7-881b-407c0222527e",
    authority: "https://login.microsoftonline.com/9152bf5c-22ff-4e4a-8624-784a2d243006",
    // Both production and localhost are registered as SPA redirect URIs.
    // Keeping this aligned with auth.js lets Graph token requests return to
    // the same origin where the app is currently running.
    redirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};

const SCOPES = ["User.Read"];
const SCOPES_GRAPH    = ["User.Read", "Files.ReadWrite", "Mail.Read", "Mail.Send", "Calendars.Read", "Tasks.ReadWrite"];
const SCOPES_DRAFTS   = ["User.Read", "Mail.Read", "Mail.ReadWrite", "Mail.Send"];
const SCOPES_TODO     = ["User.Read", "Tasks.ReadWrite"];
const SCOPES_CONTACTS = ["User.Read", "Contacts.Read", "People.Read", "User.ReadBasic.All"];
const ADMIN_EMAILS    = ["propertymanagement2@dhananipeg.com", "systemmanager1@dhananipeg.com"];
function isAdmin(){ return currentUser && ADMIN_EMAILS.includes(currentUser.email?.toLowerCase()); }
const ONEDRIVE_FOLDER = "DPEGTaskManager";

// PRINCIPALS — those with Wednesday Review vs Discussion Notes
const PRINCIPALS = {
  "nikhil@dhananipeg.com":  { name:"Nikhil Dhanani",  role:"President",                               wednesday: true  },
  "nick@dhananipeg.com":    { name:"Nick Dhanani",     role:"Chief Executive Officer",                 wednesday: false },
  "ali@dhananipeg.com":     { name:"Ali Wadhwani",     role:"Chief Financial Officer",                 wednesday: false },
  "nurali@dhananipeg.com":  { name:"Nurali Wadhwani",  role:"Chief Financial Officer",                 wednesday: false },
  "lucy@dhananipeg.com":    { name:"Lucy Singh",        role:"Chief Operating Officer",                 wednesday: false },
  "faiz@dhananipeg.com":    { name:"Faiz Hirani",       role:"Principal, Investor Relations",           wednesday: false },
  "junior@dhananipeg.com":  { name:"Junior Dhanani",    role:"Principal, Retail Assets & Acquisitions", wednesday: false },
  "rahul@dhananipeg.com":   { name:"Rahul Wadhwani",    role:"Principal, Marketing",                    wednesday: false },
};

let msalInstance = null;
let currentUser = null;
let currentAccount = null;
let isWednesdayUser = false;
let tasks = [];
let archives = [];
let staffConfig = {};
let userContacts = {};
let customDepartments = [];
let sharedDepartmentsVersion = null;
let customNotes = [];
let sharedDataActive = false;
let sharedDataVersion = null;
let selectedTaskIds = new Set();
let ntAssignees = [];
let curTaskId = null, curWeek = 0, curSearch = "", curPplSel = null, curDeptSel = null, curPplFilter = "", directoryMode = "people";
let showMasterCompleted = false;
let CH = {};
const PROOF_RETURN_KEY='dpeg_proof_return_search';

// ============================================================
// PEOPLE & DEPARTMENTS DATA
// ============================================================
const DEPARTMENTS = [
  "Investor Relations","Accounting","Acquisitions","Development","Software Development",
  "Construction","Property Management","Maintenance","Marketing",
  "Legal and Title","Leasing","IT","Operations","Lending",
  "Insurance","Multifamily","EB-5","Outside DPEG",
];

const DEPT_COLORS = {
  "Investor Relations":"#0E3416","Accounting":"#1A5C2A","Acquisitions":"#2E7D3F",
  "Development":"#1A237E","Software Development":"#1B5E20","Construction":"#BF360C","Property Management":"#006064",
  "Maintenance":"#37474F","Marketing":"#4A148C","Legal and Title":"#6D4C41",
  "Leasing":"#0D47A1","IT":"#1B5E20","Operations":"#212121",
  "Lending":"#880E4F","Insurance":"#E65100","Multifamily":"#33691E","EB-5":"#1A237E",
  "Outside DPEG":"#475569",
};

function userContactsKey(){return `dpeg_user_contacts_${normEmail(currentUser?.email||'unknown')}`;}
function loadUserContacts(){
  try{userContacts=JSON.parse(localStorage.getItem(userContactsKey())||'{}')||{};}catch{userContacts={};}
}
function saveUserContacts(){
  try{localStorage.setItem(userContactsKey(),JSON.stringify(userContacts));}catch{}
}
