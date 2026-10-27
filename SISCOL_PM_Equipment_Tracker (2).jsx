import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard,
  ClipboardCheck,
  Truck,
  FolderKanban,
  History as HistoryIcon,
  BarChart3,
  Users as UsersIcon,
  Plus,
  X,
  Search,
  AlertTriangle,
  Trash2,
  Pencil,
  RefreshCw,
  Check,
  ChevronDown,
  ChevronRight,
  Cloud,
  CloudOff,
  IndianRupee,
  Wrench,
  CircleSlash,
  PauseCircle,
  LogOut,
  Lock,
  ShieldCheck,
  UserCircle2,
  FileDown,
  ArrowLeft,
  Info,
} from "lucide-react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  LineChart,
  Line,
  Legend,
} from "recharts";

/* ==================================================================== */
/* Constants & seed data                                                */
/* ==================================================================== */

const STORAGE_KEY = "siscol_pm_tracker_v2";
const SHARED = true; // multi-site data: all registered users share one live dataset

const SCOPE_ALL = "all";
const SCOPE_SISCOL = "siscol";

const ROLE_ADMIN = "admin";
const ROLE_INCHARGE = "incharge";

const STATUS = {
  GOOD: "Good",
  BREAKDOWN: "Breakdown",
  MAINTENANCE: "Under Maintenance",
  IDLE: "Idle",
  DEHIRED: "Dehired",
};

const STATUS_ORDER = [STATUS.GOOD, STATUS.BREAKDOWN, STATUS.MAINTENANCE, STATUS.IDLE, STATUS.DEHIRED];

const STATUS_STYLE = {
  [STATUS.GOOD]: { color: "#3FBE83", bg: "rgba(63,190,131,0.13)", icon: Check, label: "Operational" },
  [STATUS.BREAKDOWN]: { color: "#E5544B", bg: "rgba(229,84,75,0.14)", icon: AlertTriangle, label: "Breakdown" },
  [STATUS.MAINTENANCE]: { color: "#E3A438", bg: "rgba(227,164,56,0.14)", icon: Wrench, label: "Under Maintenance" },
  [STATUS.IDLE]: { color: "#8B93A1", bg: "rgba(139,147,161,0.14)", icon: PauseCircle, label: "Idle" },
  [STATUS.DEHIRED]: { color: "#767585", bg: "rgba(118,117,133,0.14)", icon: CircleSlash, label: "Dehired" },
};

const EQUIPMENT_TYPES = [
  "Boom Lift 50ft", "Boom Lift 60ft", "Boom Lift 65ft", "Boom Lift 66ft",
  "Boom Lift 80ft", "Boom Lift 90ft", "Crane 25T", "Crane 30T", "Crane 60T",
  "Crane 70T", "Deck Rider 15m", "Farahna 15T", "Farahna 21T", "Farahna 23T",
  "Farahna 25T", "Hydra 15T", "Hydra 20T", "Trailer", "Other",
];

const SEED_EQUIPMENT = [{"id": "EQ001", "project": "C5", "type": "Trailer", "vendor": "ANIL", "rentedBy": "SISCOL", "shifts": 2, "rate": 123000, "model": "TATA MOTORS LTD", "yom": 2011, "vehicleNo": "GJ00601820000182", "contact": "8825413565", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ002", "project": "C5", "type": "Trailer", "vendor": "ANIL", "rentedBy": "SISCOL", "shifts": 2, "rate": 123000, "model": "TATA MOTORS LTD", "yom": 2008, "vehicleNo": "GJ00600280009442", "contact": "8825413565", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ003", "project": "C5", "type": "Farahna 23T", "vendor": "BHUPENDAR", "rentedBy": "SISCOL", "shifts": 2, "rate": 205000, "model": "ESCORTS F23", "yom": 2026, "vehicleNo": "ECE02205VT0123213", "contact": "9512555777", "tpiExpiry": "2027-01-13", "status": "Good", "active": true}, {"id": "EQ004", "project": "C5", "type": "Farahna 15T", "vendor": "SHAIL", "rentedBy": "SISCOL", "shifts": 2, "rate": 205000, "model": "ACE FX150", "yom": 2024, "vehicleNo": "ACEFX150TG0438070", "contact": "9737980487", "tpiExpiry": "2027-02-10", "status": "Good", "active": true}, {"id": "EQ005", "project": "CCS", "type": "Trailer", "vendor": "Datar services Pvt Ltd", "rentedBy": "SISCOL", "shifts": 2, "rate": 145000, "model": "ASHOK LEYLAND TRAILER", "yom": 2021, "vehicleNo": "HR 38 AB 3522", "contact": "9818082784", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ006", "project": "CCS", "type": "Trailer", "vendor": "Datar services Pvt Ltd", "rentedBy": "SISCOL", "shifts": 2, "rate": 145000, "model": "TATA SIGMA", "yom": 2026, "vehicleNo": "HR 38 AL 0923", "contact": "9818082784", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ007", "project": "CCS", "type": "Deck Rider 15m", "vendor": "NSS-Chennai", "rentedBy": "SISCOL", "shifts": 2, "rate": 220000, "model": "DECK RIDER", "yom": 2014, "vehicleNo": "QRA-10377", "contact": "9161520571", "tpiExpiry": "2027-02-03", "status": "Good", "active": true}, {"id": "EQ008", "project": "CCS", "type": "Deck Rider 15m", "vendor": "NSS-Chennai", "rentedBy": "SISCOL", "shifts": 2, "rate": 220000, "model": "DECK RIDER", "yom": 2016, "vehicleNo": "QRA-11749", "contact": "9161520571", "tpiExpiry": "2026-11-13", "status": "Good", "active": true}, {"id": "EQ009", "project": "CCS", "type": "Hydra 15T", "vendor": "SAI cranes(Mr.Rajesh)", "rentedBy": "SISCOL", "shifts": 1, "rate": 95000, "model": "ACE F150", "yom": 2023, "vehicleNo": "HR 51 CL 6735", "contact": "9250019091", "tpiExpiry": "2026-09-08", "status": "Good", "active": true}, {"id": "EQ010", "project": "CCS", "type": "Hydra 20T", "vendor": "SAI cranes(Mr.Rajesh)", "rentedBy": "SISCOL", "shifts": 2, "rate": 200000, "model": "ACE F210", "yom": 2022, "vehicleNo": "HR 51 C G7462", "contact": "9250019091", "tpiExpiry": "2027-02-03", "status": "Good", "active": true}, {"id": "EQ011", "project": "Chalet", "type": "Farahna 15T", "vendor": "Amruth Enterprises", "rentedBy": "SISCOL", "shifts": 1, "rate": 120000, "model": "F150", "yom": 2024, "vehicleNo": "ACE0F150TT06", "contact": "9323583186", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ012", "project": "Chalet", "type": "Farahna 15T", "vendor": "Amruth Enterprises", "rentedBy": "SISCOL", "shifts": 1, "rate": 120000, "model": "F150", "yom": 2026, "vehicleNo": "ACE0F150TR0476251", "contact": "9323583186", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ013", "project": "Chalet", "type": "Deck Rider 15m", "vendor": "Irise India", "rentedBy": "SISCOL", "shifts": 1, "rate": 220000, "model": "Make: Nifty, Model No.: HR 15D,", "yom": 2021, "vehicleNo": "TN92-F0235", "contact": "8652739212", "tpiExpiry": "2026-12-15", "status": "Good", "active": true}, {"id": "EQ014", "project": "Chalet", "type": "Deck Rider 15m", "vendor": "Irise India", "rentedBy": "SISCOL", "shifts": 1, "rate": 220000, "model": "Make: Nifty, Model No.: HR 15 D,Sr", "yom": null, "vehicleNo": "DXT809513", "contact": "8652739212", "tpiExpiry": "2026-12-16", "status": "Good", "active": true}, {"id": "EQ015", "project": "Chalet", "type": "Deck Rider 15m", "vendor": "NSS", "rentedBy": "SISCOL", "shifts": 2, "rate": 340000, "model": "Make: NIFTY, Model No.: HR 15,Sr No.: 09", "yom": 2025, "vehicleNo": "B300027040", "contact": "8081355656", "tpiExpiry": "2026-12-17", "status": "Good", "active": true}, {"id": "EQ016", "project": "Chalet", "type": "Deck Rider 15m", "vendor": "NSS", "rentedBy": "SISCOL", "shifts": 2, "rate": 340000, "model": "Make: NIFTY, Model No.: DR15D,Sr No.: 09", "yom": 2026, "vehicleNo": "B300031334", "contact": "9350049559", "tpiExpiry": "2026-11-13", "status": "Good", "active": true}, {"id": "EQ017", "project": "Chalet", "type": "Trailer", "vendor": "VIJAY CARGO MOVERS", "rentedBy": "SISCOL", "shifts": 2, "rate": 110000, "model": null, "yom": null, "vehicleNo": "", "contact": "9321451545", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ018", "project": "Chalet", "type": "Trailer", "vendor": "VIJAY CARGO MOVERS", "rentedBy": "SISCOL", "shifts": 2, "rate": 110000, "model": null, "yom": null, "vehicleNo": "", "contact": "9321451545", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ019", "project": "D2 Sabarmati", "type": "Trailer", "vendor": "Agarwal Transport Service", "rentedBy": "SISCOL", "shifts": 2, "rate": 140000, "model": "TAILER 40 FEET", "yom": 2017, "vehicleNo": "RJ14GJ1199", "contact": "9414002962", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ020", "project": "D2 Sabarmati", "type": "Boom Lift 66ft", "vendor": "CRESENT POWER ENTERPRISES", "rentedBy": "SISCOL", "shifts": 2, "rate": 170000, "model": "BOOM LIFT 66 FEET JLG 600S", "yom": 2012, "vehicleNo": "300063743", "contact": "7021933268", "tpiExpiry": "2026-11-28", "status": "Good", "active": true}, {"id": "EQ021", "project": "D2 Sabarmati", "type": "Boom Lift 65ft", "vendor": "CRESENT POWER ENTERPRISES", "rentedBy": "SISCOL", "shifts": 2, "rate": 170000, "model": "BOOM LIFT 65 FEET GENIE", "yom": 2021, "vehicleNo": "300075479", "contact": "7021933268", "tpiExpiry": "2026-12-01", "status": "Good", "active": true}, {"id": "EQ022", "project": "D2 Sabarmati", "type": "Crane 30T", "vendor": "Gurunanak Crane Service", "rentedBy": "SISCOL", "shifts": 1, "rate": 150000, "model": "30 TON CRANE", "yom": 2025, "vehicleNo": "GJ27CQ2140", "contact": "9727044111", "tpiExpiry": "2027-02-11", "status": "Good", "active": true}, {"id": "EQ023", "project": "D2 Sabarmati", "type": "Farahna 15T", "vendor": "Gurunanak Crane Service", "rentedBy": "SISCOL", "shifts": 1, "rate": 110000, "model": "F-15 ESCORTS", "yom": 2022, "vehicleNo": "GJ27CQ1421", "contact": "9727044111", "tpiExpiry": "2027-02-11", "status": "Good", "active": true}, {"id": "EQ024", "project": "D2 Sabarmati", "type": "Boom Lift 66ft", "vendor": "L&T", "rentedBy": "L&T", "shifts": 2, "rate": 170000, "model": "JLG 66 FEET BOOMLIFT", "yom": "L&T", "vehicleNo": "L&T", "contact": "9381906734", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ025", "project": "D2 Sabarmati", "type": "Boom Lift 66ft", "vendor": "L&T", "rentedBy": "L&T", "shifts": 2, "rate": 170000, "model": "JLG 66 FEET BOOMLIFT", "yom": "L&T", "vehicleNo": "L&T", "contact": "9381906734", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ026", "project": "D2 Sabarmati", "type": "Boom Lift 90ft", "vendor": "Manlift", "rentedBy": "SISCOL", "shifts": 2, "rate": 205000, "model": "BOOM LIFT 90 FEET JLG 860SJ", "yom": 2022, "vehicleNo": "B300002822", "contact": "8826071117", "tpiExpiry": "2027-01-17", "status": "Good", "active": true}, {"id": "EQ027", "project": "D2 Sabarmati", "type": "Boom Lift 90ft", "vendor": "Manlift", "rentedBy": "SISCOL", "shifts": 2, "rate": 205000, "model": "BOOM LIFT 90 FEET JLG", "yom": 2021, "vehicleNo": "B300017640", "contact": "8826071117", "tpiExpiry": "2027-02-11", "status": "Good", "active": true}, {"id": "EQ028", "project": "D2 Sabarmati", "type": "Boom Lift 90ft", "vendor": "Navya Crane Service", "rentedBy": "SISCOL", "shifts": 2, "rate": 205000, "model": "BOOM LIFT 90 FEET JLG 860SJ", "yom": 2023, "vehicleNo": "300193576", "contact": "7906076950", "tpiExpiry": "2026-11-15", "status": "Good", "active": true}, {"id": "EQ029", "project": "D2 Sabarmati", "type": "Boom Lift 90ft", "vendor": "Navya Crane Service", "rentedBy": "SISCOL", "shifts": 1, "rate": 160000, "model": "JLG860SJ 90 FEET", "yom": 2024, "vehicleNo": "300128435", "contact": "7906076950", "tpiExpiry": "2027-02-21", "status": "Good", "active": true}, {"id": "EQ030", "project": "D2 Sabarmati", "type": "Boom Lift 60ft", "vendor": "Navya Crane Service", "rentedBy": "SISCOL", "shifts": 2, "rate": 175000, "model": "Boom Lifter 60 Feet", "yom": 2014, "vehicleNo": "B300000620", "contact": "7906076950", "tpiExpiry": "2027-01-14", "status": "Good", "active": true}, {"id": "EQ031", "project": "D2 Sabarmati", "type": "Boom Lift 66ft", "vendor": "Navya Crane Service", "rentedBy": "SISCOL", "shifts": 2, "rate": 175000, "model": "Boom Lifter 66 Feet", "yom": 2023, "vehicleNo": "300118686", "contact": "7906076950", "tpiExpiry": "2026-09-20", "status": "Good", "active": true}, {"id": "EQ032", "project": "D2 Sabarmati", "type": "Boom Lift 90ft", "vendor": "Reach International", "rentedBy": "SISCOL", "shifts": 2, "rate": 205000, "model": "BOOM LIFT 90 FEET JLG", "yom": 2019, "vehicleNo": "B3000028269", "contact": "8800040314", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ033", "project": "D2 Sabarmati", "type": "Crane 70T", "vendor": "Satyam Heavy Lifters", "rentedBy": "SISCOL", "shifts": 2, "rate": 450000, "model": "70 TON CRANE - XCMG", "yom": 2024, "vehicleNo": "MH01ER4327", "contact": "9512555777", "tpiExpiry": "2026-09-08", "status": "Good", "active": true}, {"id": "EQ034", "project": "D2 Sabarmati", "type": "Crane 60T", "vendor": "Satyam Heavy Lifters", "rentedBy": "SISCOL", "shifts": 2, "rate": 390000, "model": "60 TON CRANE", "yom": 2022, "vehicleNo": "GJ18SS1380", "contact": "9512555777", "tpiExpiry": "2026-12-25", "status": "Good", "active": true}, {"id": "EQ035", "project": "D2 Sabarmati", "type": "Farahna 23T", "vendor": "Satyam Heavy Lifters", "rentedBy": "SISCOL", "shifts": 2, "rate": 225000, "model": "F-23", "yom": 2025, "vehicleNo": "GJ27CQ3534", "contact": "9512555777", "tpiExpiry": "2026-12-29", "status": "Good", "active": true}, {"id": "EQ036", "project": "D2 Sabarmati", "type": "Farahna 15T", "vendor": "Satyam Heavy Lifters", "rentedBy": "SISCOL", "shifts": 2, "rate": 180000, "model": "ACE-150", "yom": 2025, "vehicleNo": "GJ27CQ3594", "contact": "9512555777", "tpiExpiry": "2027-01-29", "status": "Good", "active": true}, {"id": "EQ037", "project": "D2 Sabarmati", "type": "Boom Lift 60ft", "vendor": "Sudhir Power limted", "rentedBy": "SISCOL", "shifts": 2, "rate": 175000, "model": "Boom Lifter 60 Feet", "yom": 2016, "vehicleNo": "S60D620", "contact": "9911598806", "tpiExpiry": "2026-09-05", "status": "Breakdown", "active": true}, {"id": "EQ038", "project": "Senapathi", "type": "Boom Lift 60ft", "vendor": "SISCOL", "rentedBy": "SISCOL", "shifts": 1, "rate": 120000, "model": "60ft - JLG600S", "yom": 2019, "vehicleNo": "", "contact": "9766166940", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ039", "project": "Senapathi", "type": "Boom Lift 60ft", "vendor": "Premier Infra", "rentedBy": "Premier Infra", "shifts": 1, "rate": 120000, "model": "60ft - JLG600S", "yom": 2017, "vehicleNo": "Premier Infra", "contact": "7989638603", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ040", "project": "Senapathi", "type": "Farahna 15T", "vendor": "Premier Infra", "rentedBy": "Premier Infra", "shifts": 1, "rate": 130000, "model": "Escort 15T-C7335571", "yom": 2022, "vehicleNo": "Premier Infra", "contact": "7989638603", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ041", "project": "Senapathi", "type": "Farahna 23T", "vendor": "Premier Infra", "rentedBy": "Premier Infra", "shifts": 1, "rate": 23000, "model": "Escort 23T - Daily rental", "yom": 2021, "vehicleNo": "Premier Infra", "contact": "7989638603", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ042", "project": "Senapathi", "type": "Farahna 15T", "vendor": "Galaxy", "rentedBy": "Galaxy", "shifts": 1, "rate": 12500, "model": "Escort 15T - Daily rental", "yom": 2022, "vehicleNo": "Galaxy", "contact": "9322452985", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ043", "project": "Senapathi", "type": "Boom Lift 60ft", "vendor": "Galaxy", "rentedBy": "Galaxy", "shifts": 1, "rate": 120000, "model": "60ft - JLG600S", "yom": 2020, "vehicleNo": "Galaxy", "contact": "9322452985", "tpiExpiry": null, "status": "Good", "active": true}, {"id": "EQ044", "project": "Vikram Solar", "type": "Boom Lift 60ft", "vendor": "GMMCO", "rentedBy": "Star", "shifts": 1, "rate": 130000, "model": "JLG660SJ", "yom": 2025, "vehicleNo": "B300027040", "contact": "9094675905", "tpiExpiry": "2027-03-26", "status": "Good", "active": true}, {"id": "EQ045", "project": "Vikram Solar", "type": "Boom Lift 60ft", "vendor": "GMMCO", "rentedBy": "Star", "shifts": 1, "rate": 130000, "model": "JLGM660SJ", "yom": 2026, "vehicleNo": "B300031334", "contact": "9094675905", "tpiExpiry": "2027-04-15", "status": "Good", "active": true}, {"id": "EQ046", "project": "Vikram Solar", "type": "Boom Lift 60ft", "vendor": "GMMCO", "rentedBy": "Star", "shifts": 1, "rate": 130000, "model": "JLG660SJ", "yom": 2026, "vehicleNo": "B300031332", "contact": "9094675905", "tpiExpiry": "2027-05-14", "status": "Good", "active": true}, {"id": "EQ047", "project": "Vikram Solar", "type": "Boom Lift 50ft", "vendor": "JP", "rentedBy": "Star", "shifts": 1, "rate": 130000, "model": "JLG600S", "yom": 2020, "vehicleNo": "B300031102", "contact": "9840497366", "tpiExpiry": "2026-12-01", "status": "Good", "active": true}, {"id": "EQ048", "project": "Vikram Solar", "type": "Crane 25T", "vendor": "MK", "rentedBy": "Star", "shifts": 1, "rate": 210000, "model": "QY-25K5", "yom": 2018, "vehicleNo": "MH01-EF9352", "contact": "6381217889", "tpiExpiry": "2028-04-10", "status": "Breakdown", "active": true}, {"id": "EQ049", "project": "Vikram Solar", "type": "Boom Lift 60ft", "vendor": "MT&T", "rentedBy": "SISCOL", "shifts": 2, "rate": 175000, "model": "JLG660SJ", "yom": 2021, "vehicleNo": "B300009190", "contact": "8700576665/7339689952", "tpiExpiry": "2026-11-18", "status": "Good", "active": true}, {"id": "EQ050", "project": "Vikram Solar", "type": "Boom Lift 60ft", "vendor": "MT&T", "rentedBy": "SISCOL", "shifts": 2, "rate": 175000, "model": "660SJ", "yom": 2021, "vehicleNo": "B300004509", "contact": "8700576665/7339689952", "tpiExpiry": "2026-12-29", "status": "Good", "active": true}, {"id": "EQ051", "project": "Vikram Solar", "type": "Boom Lift 60ft", "vendor": "Global", "rentedBy": "Nihal", "shifts": 2, "rate": 120000, "model": "JLG660AJ", "yom": 2020, "vehicleNo": "300154297", "contact": "9513456705", "tpiExpiry": "2026-12-22", "status": "Good", "active": true}, {"id": "EQ052", "project": "Vikram Solar", "type": "Farahna 21T", "vendor": "VP", "rentedBy": "Nihal", "shifts": 1, "rate": 150000, "model": "Farhana-21", "yom": 2020, "vehicleNo": "TN92-F0235", "contact": "9994732762", "tpiExpiry": "2027-01-01", "status": "Good", "active": true}, {"id": "EQ053", "project": "Vikram Solar", "type": "Crane 25T", "vendor": "Urmila", "rentedBy": "SISCOL", "shifts": 1, "rate": 240000, "model": "XCMG/XCT25L4Y", "yom": 2024, "vehicleNo": "MH02-GJ3043", "contact": "9962665999", "tpiExpiry": "2027-06-08", "status": "Good", "active": true}, {"id": "EQ054", "project": "Vikram Solar", "type": "Crane 30T", "vendor": "Srivaari", "rentedBy": "Nihal", "shifts": 1, "rate": 225000, "model": "ZTC300V", "yom": 2021, "vehicleNo": "GJ27-CQ2775", "contact": "9843870170", "tpiExpiry": "2026-12-18", "status": "Good", "active": true}, {"id": "EQ055", "project": "Vikram Solar", "type": "Farahna 25T", "vendor": "Suresh Cranes", "rentedBy": "Nihal", "shifts": 1, "rate": 170000, "model": "Farhana-25", "yom": 2026, "vehicleNo": "DXT809513", "contact": "9751879945", "tpiExpiry": "2027-06-08", "status": "Good", "active": true}, {"id": "EQ056", "project": "Vikram Solar", "type": "Boom Lift - 60ft", "vendor": "MT&T - 60ft", "rentedBy": "SISCOL", "shifts": 1, "rate": 130000, "model": "JLG660SJ", "yom": 2021, "vehicleNo": "B300009153", "contact": "8700576665/7339689952", "tpiExpiry": "2026-11-19", "status": "Good", "active": true}, {"id": "EQ057", "project": "Vikram Solar", "type": "Boom Lift - 60ft", "vendor": "MT&T - 60ft", "rentedBy": "SISCOL", "shifts": 2, "rate": 175000, "model": "660SJ", "yom": 2021, "vehicleNo": "B300009039", "contact": "8700576665/7339689952", "tpiExpiry": "2026-12-28", "status": "Good", "active": true}, {"id": "EQ058", "project": "Vikram Solar", "type": "Boom Lift 80ft", "vendor": "MT&T - 80ft", "rentedBy": "SISCOL", "shifts": 1, "rate": 180000, "model": "BT28RT", "yom": 2023, "vehicleNo": "BT20255", "contact": "8700576665/7339689952", "tpiExpiry": "2027-01-26", "status": "Good", "active": true}];

const SEED_PROJECTS = Array.from(new Set(SEED_EQUIPMENT.map((e) => e.project))).sort();

function seedData() {
  return {
    projects: SEED_PROJECTS,
    equipment: SEED_EQUIPMENT,
    entries: [],
    users: [],
    meta: { lastEqNum: SEED_EQUIPMENT.length },
  };
}

/* ==================================================================== */
/* Helpers                                                              */
/* ==================================================================== */

function localDateStr(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function todayStr() {
  return localDateStr(new Date());
}

function fmtDateShort(iso) {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}

function fmtDateTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleString("en-IN", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });
}

function daysFromToday(iso) {
  if (!iso) return null;
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d.getTime())) return null;
  const t = new Date(todayStr() + "T00:00:00");
  return Math.round((d.getTime() - t.getTime()) / 86400000);
}

function fmtINR(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return "₹" + Number(n).toLocaleString("en-IN");
}

function uid(prefix) {
  return prefix + Math.random().toString(36).slice(2, 9);
}

function classNames(...xs) {
  return xs.filter(Boolean).join(" ");
}

function exportToExcel(data, latestEntryByEquipment, scopeLabel) {
  const wb = XLSX.utils.book_new();

  const eqRows = data.equipment.map((e) => {
    const le = latestEntryByEquipment[e.id];
    return {
      "Equipment ID": e.id,
      "Project": e.project,
      "Type": e.type,
      "Vendor": e.vendor,
      "Rented By": e.rentedBy || "SISCOL",
      "Vehicle / Serial No.": e.vehicleNo || "",
      "Model": e.model || "",
      "Year of Mfg": e.yom || "",
      "Rate (INR)": e.rate || 0,
      "Shifts": e.shifts || "",
      "Contact": e.contact || "",
      "TPI Expiry": e.tpiExpiry || "",
      "Current Status": e.status,
      "Last Logged Date": le ? le.date : "",
      "Last Logged By": le ? le.loggedBy || "" : "",
    };
  });
  const eqSheet = XLSX.utils.json_to_sheet(eqRows);
  XLSX.utils.book_append_sheet(wb, eqSheet, "Equipment");

  const eqById = {};
  data.equipment.forEach((e) => (eqById[e.id] = e));
  const historyRows = data.entries
    .slice()
    .sort((a, b) => (a.date < b.date ? 1 : -1))
    .map((en) => {
      const e = eqById[en.equipmentId] || {};
      return {
        "Date": en.date,
        "Project": e.project || "",
        "Equipment": e.type || "",
        "Equipment ID": en.equipmentId,
        "Vehicle No.": e.vehicleNo || "",
        "Rented By": e.rentedBy || "SISCOL",
        "Status": en.status,
        "Reason": en.reason || "",
        "Remark": en.remark || "",
        "Logged By": en.loggedBy || "",
        "Logged At": en.loggedAt ? fmtDateTime(en.loggedAt) : "",
      };
    });
  const historySheet = XLSX.utils.json_to_sheet(historyRows);
  XLSX.utils.book_append_sheet(wb, historySheet, "Daily Log");

  const projectRows = data.projects.map((p) => {
    const eqs = data.equipment.filter((e) => e.project === p);
    return {
      "Project": p,
      "Equipment Count": eqs.length,
      "Total Rent (INR)": eqs.reduce((s, e) => s + (Number(e.rate) || 0), 0),
      "Operational": eqs.filter((e) => e.status === STATUS.GOOD).length,
      "Breakdown": eqs.filter((e) => e.status === STATUS.BREAKDOWN).length,
      "Under Maintenance": eqs.filter((e) => e.status === STATUS.MAINTENANCE).length,
      "Idle": eqs.filter((e) => e.status === STATUS.IDLE).length,
      "Dehired": eqs.filter((e) => e.status === STATUS.DEHIRED).length,
    };
  });
  const projectSheet = XLSX.utils.json_to_sheet(projectRows);
  XLSX.utils.book_append_sheet(wb, projectSheet, "Project Summary");

  XLSX.writeFile(wb, `SISCOL_PM_Export_${scopeLabel}_${todayStr()}.xlsx`);
}

/* ==================================================================== */
/* Logo                                                                 */
/* ==================================================================== */

function Logo({ size = 34 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="scLogoGrad" x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
          <stop stopColor="#FF7A45" />
          <stop offset="1" stopColor="#E5544B" />
        </linearGradient>
      </defs>
      <path d="M24 2L44 12V28C44 38 35.5 44.5 24 46C12.5 44.5 4 38 4 28V12L24 2Z" fill="url(#scLogoGrad)" />
      <path
        d="M18 30C18 32.7614 20.6863 35 24 35C27.3137 35 30 32.7614 30 30C30 27.4 27.8 26 24.5 25.2C21.2 24.4 19 23 19 20.5C19 18.2386 21.2386 16.4 24 16.4C26.4 16.4 28.4 17.7 29.2 19.6"
        stroke="#14171B"
        strokeWidth="3"
        strokeLinecap="round"
      />
    </svg>
  );
}

/* ==================================================================== */
/* Small reusable UI                                                    */
/* ==================================================================== */

function StatusBadge({ status, size = "md" }) {
  const s = STATUS_STYLE[status] || STATUS_STYLE[STATUS.GOOD];
  const Icon = s.icon;
  return (
    <span className={classNames("sc-badge", size === "sm" && "sc-badge-sm")} style={{ color: s.color, background: s.bg }}>
      <Icon size={size === "sm" ? 11 : 13} strokeWidth={2.4} />
      {s.label}
    </span>
  );
}

function StatCard({ label, value, sub, accent, icon: Icon, onClick }) {
  const Tag = onClick ? "button" : "div";
  return (
    <Tag className={classNames("sc-stat", onClick && "sc-stat-clickable")} onClick={onClick}>
      <div className="sc-stat-top">
        <span className="sc-stat-label">{label}</span>
        {Icon && (
          <span className="sc-stat-icon" style={{ background: accent ? `${accent}22` : "rgba(139,147,161,0.14)" }}>
            <Icon size={15} color={accent || "#8B93A1"} strokeWidth={2.2} />
          </span>
        )}
      </div>
      <div className="sc-stat-value" style={accent ? { color: accent } : undefined}>{value}</div>
      {sub && <div className="sc-stat-sub">{sub}</div>}
      {onClick && <span className="sc-stat-cta">View details <ChevronRight size={12} /></span>}
    </Tag>
  );
}

function Modal({ title, subtitle, onClose, children, wide }) {
  return (
    <div className="sc-modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={classNames("sc-modal", wide && "sc-modal-wide")}>
        <div className="sc-modal-head">
          <div>
            <h3>{title}</h3>
            {subtitle && <p>{subtitle}</p>}
          </div>
          <button className="sc-icon-btn" onClick={onClose}><X size={18} /></button>
        </div>
        <div className="sc-modal-body">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children, hint, group }) {
  const Tag = group ? "div" : "label";
  return (
    <Tag className="sc-field" role={group ? "group" : undefined} aria-label={group ? label : undefined}>
      <span className="sc-field-label">{label}</span>
      {children}
      {hint && <span className="sc-field-hint">{hint}</span>}
    </Tag>
  );
}

function Select({ value, onChange, options, disabled }) {
  return (
    <div className={classNames("sc-select", disabled && "sc-select-disabled")}>
      <select value={value} disabled={disabled} onChange={(e) => onChange(e.target.value)}>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
      <ChevronDown size={14} />
    </div>
  );
}

function PageHead({ title, sub, action }) {
  return (
    <div className="sc-page-head">
      <div>
        <h1>{title}</h1>
        {sub && <p>{sub}</p>}
      </div>
      {action}
    </div>
  );
}

/* ==================================================================== */
/* Auth — lightweight sign-in gate                                      */
/* ==================================================================== */

function AuthGate({ users, projects, onRegister, onLogin }) {
  const [mode, setMode] = useState(users.length === 0 ? "register" : "login");
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [role, setRole] = useState(ROLE_INCHARGE);
  const [project, setProject] = useState(projects[0] || "");
  const [selectedUser, setSelectedUser] = useState(users[0]?.id || "");
  const [error, setError] = useState("");

  function submitRegister() {
    const clean = name.trim();
    if (!clean) { setError("Enter your full name."); return; }
    if (!/^\d{4,6}$/.test(pin)) { setError("PIN must be 4–6 digits."); return; }
    if (users.some((u) => u.name.toLowerCase() === clean.toLowerCase())) { setError("That name is already registered — use Sign in instead."); return; }
    if (role === ROLE_INCHARGE && !project) { setError("Select the site you're in charge of."); return; }
    onRegister({ name: clean, pin, role, project: role === ROLE_INCHARGE ? project : null });
  }

  function submitLogin() {
    const u = users.find((x) => x.id === selectedUser);
    if (!u) { setError("Select your name."); return; }
    if (u.pin !== pin) { setError("Incorrect PIN."); return; }
    onLogin(u);
  }

  return (
    <div className="sc-auth-root">
      <div className="sc-auth-card">
        <div className="sc-auth-brand">
          <Logo size={44} />
          <div>
            <div className="sc-auth-brand-name">SISCOL</div>
            <div className="sc-auth-brand-sub">Plant &amp; Machinery Tracker</div>
          </div>
        </div>

        <div className="sc-segmented sc-auth-tabs">
          <button className={classNames("sc-segmented-opt", mode === "login" && "sc-segmented-opt-active")} onClick={() => { setMode("login"); setError(""); }}>
            Sign in
          </button>
          <button className={classNames("sc-segmented-opt", mode === "register" && "sc-segmented-opt-active")} onClick={() => { setMode("register"); setError(""); }}>
            Register
          </button>
        </div>

        {mode === "login" ? (
          <div className="sc-form">
            {users.length === 0 ? (
              <div className="sc-empty">No one is registered yet — switch to Register.</div>
            ) : (
              <>
                <Field label="Your name">
                  <Select value={selectedUser} onChange={setSelectedUser} options={users.map((u) => u.id)} />
                  <span className="sc-field-hint">{users.find((u) => u.id === selectedUser)?.name}{" — "}{users.find((u) => u.id === selectedUser)?.role === ROLE_ADMIN ? "Admin" : `Site incharge, ${users.find((u) => u.id === selectedUser)?.project}`}</span>
                </Field>
                <Field label="PIN">
                  <input className="sc-input" type="password" inputMode="numeric" maxLength={6} value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} onKeyDown={(e) => e.key === "Enter" && submitLogin()} />
                </Field>
                {error && <div className="sc-error">{error}</div>}
                <button className="sc-btn sc-btn-primary sc-btn-block" onClick={submitLogin}><Lock size={14} /> Sign in</button>
              </>
            )}
          </div>
        ) : (
          <div className="sc-form">
            <Field label="Full name">
              <input className="sc-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Ramesh Kumar" />
            </Field>
            <Field label="Role" group>
              <div className="sc-status-picker">
                <button type="button" className={classNames("sc-status-opt", role === ROLE_INCHARGE && "sc-status-opt-active")} onClick={() => setRole(ROLE_INCHARGE)}>Site incharge</button>
                <button type="button" className={classNames("sc-status-opt", role === ROLE_ADMIN && "sc-status-opt-active")} onClick={() => setRole(ROLE_ADMIN)}>Admin (HQ)</button>
              </div>
            </Field>
            {role === ROLE_INCHARGE && (
              <Field label="Site you're in charge of">
                <Select value={project} onChange={setProject} options={projects} />
              </Field>
            )}
            <Field label="Create a 4–6 digit PIN" hint="You'll use this to sign in next time — this is a lightweight access PIN, not a secure password, so don't reuse one you use elsewhere.">
              <input className="sc-input" type="password" inputMode="numeric" maxLength={6} value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} onKeyDown={(e) => e.key === "Enter" && submitRegister()} />
            </Field>
            {error && <div className="sc-error">{error}</div>}
            <button className="sc-btn sc-btn-primary sc-btn-block" onClick={submitRegister}><UserCircle2 size={15} /> Register &amp; continue</button>
          </div>
        )}

        <div className="sc-auth-foot">
          <Info size={12} />
          <span>All registered users share the same live equipment data. You'll need to sign in again each time you reopen this app.</span>
        </div>
      </div>
    </div>
  );
}

/* ==================================================================== */
/* Main App                                                             */
/* ==================================================================== */

export default function App() {
  const [data, setData] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [syncState, setSyncState] = useState("idle");
  const [toast, setToast] = useState(null);
  const [scope, setScope] = useState(SCOPE_SISCOL);
  const [currentUser, setCurrentUser] = useState(null); // session only, never persisted
  const [detailFilter, setDetailFilter] = useState(null); // status string, opens EquipmentDetailModal
  const [openProject, setOpenProject] = useState(null); // project name, opens ProjectDetailModal
  const saveTimer = useRef(null);
  const firstLoad = useRef(true);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, SHARED);
        if (cancelled) return;
        if (res && res.value) setData(normalizeData(JSON.parse(res.value)));
        else setData(seedData());
      } catch (e) {
        setData(seedData());
      } finally {
        if (!cancelled) setLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, []);

  function normalizeData(d) {
    return {
      projects: d.projects || [],
      equipment: d.equipment || [],
      entries: d.entries || [],
      users: d.users || [],
      meta: d.meta || { lastEqNum: (d.equipment || []).length },
    };
  }

  useEffect(() => {
    if (!loaded || !data) return;
    if (firstLoad.current) { firstLoad.current = false; return; }
    setSyncState("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await window.storage.set(STORAGE_KEY, JSON.stringify(data), SHARED);
        setSyncState("saved");
      } catch (e) {
        setSyncState("error");
      }
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded]);

  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast((cur) => (cur === msg ? null : cur)), 2600);
  }

  /* ---------------- mutations ---------------- */

  const registerUser = useCallback((u) => {
    const rec = { id: uid("U"), createdAt: new Date().toISOString(), ...u };
    setData((d) => ({ ...d, users: [...d.users, rec] }));
    setCurrentUser(rec);
  }, []);

  const addProject = useCallback((name) => {
    const clean = name.trim();
    if (!clean) return false;
    let ok = true;
    setData((d) => {
      if (d.projects.some((p) => p.toLowerCase() === clean.toLowerCase())) { ok = false; return d; }
      return { ...d, projects: [...d.projects, clean] };
    });
    return ok;
  }, []);

  const deleteProject = useCallback((name) => {
    let ok = true;
    setData((d) => {
      if (d.equipment.some((e) => e.project === name)) { ok = false; return d; }
      return { ...d, projects: d.projects.filter((p) => p !== name) };
    });
    return ok;
  }, []);

  const addEquipment = useCallback((eq) => {
    setData((d) => {
      const num = (d.meta.lastEqNum || d.equipment.length) + 1;
      const id = "EQ" + String(num).padStart(3, "0");
      const newEq = { id, active: true, status: STATUS.GOOD, rentedBy: "SISCOL", ...eq };
      return { ...d, equipment: [...d.equipment, newEq], meta: { ...d.meta, lastEqNum: num } };
    });
  }, []);

  const updateEquipment = useCallback((id, patch) => {
    setData((d) => ({ ...d, equipment: d.equipment.map((e) => (e.id === id ? { ...e, ...patch } : e)) }));
  }, []);

  const deleteEquipment = useCallback((id) => {
    setData((d) => ({
      ...d,
      equipment: d.equipment.filter((e) => e.id !== id),
      entries: d.entries.filter((en) => en.equipmentId !== id),
    }));
  }, []);

  const logStatus = useCallback((equipmentId, entry, loggedBy) => {
    setData((d) => {
      const existingIdx = d.entries.findIndex((e) => e.equipmentId === equipmentId && e.date === entry.date);
      const record = {
        id: existingIdx >= 0 ? d.entries[existingIdx].id : uid("EN"),
        equipmentId,
        date: entry.date,
        status: entry.status,
        reason: entry.reason || "",
        remark: entry.remark || "",
        loggedBy: loggedBy?.name || "Unknown",
        loggedByRole: loggedBy?.role || "",
        loggedAt: new Date().toISOString(),
      };
      let entries;
      if (existingIdx >= 0) { entries = d.entries.slice(); entries[existingIdx] = record; }
      else entries = [...d.entries, record];
      const eqEntries = entries.filter((e) => e.equipmentId === equipmentId).sort((a, b) => a.date.localeCompare(b.date));
      const latest = eqEntries[eqEntries.length - 1];
      const equipment = d.equipment.map((e) => (e.id === equipmentId ? { ...e, status: latest ? latest.status : e.status } : e));
      return { ...d, entries, equipment };
    });
  }, []);

  /* ---------------- derived data ---------------- */

  const rentedByOptions = useMemo(() => {
    if (!data) return ["SISCOL"];
    return Array.from(new Set(["SISCOL", ...data.equipment.map((e) => e.rentedBy || "SISCOL")]));
  }, [data]);

  const scopedEquipment = useMemo(() => {
    if (!data) return [];
    return scope === SCOPE_ALL ? data.equipment : data.equipment.filter((e) => (e.rentedBy || "SISCOL") === "SISCOL");
  }, [data, scope]);

  const scopedIds = useMemo(() => new Set(scopedEquipment.map((e) => e.id)), [scopedEquipment]);

  const scopedData = useMemo(() => {
    if (!data) return data;
    return { ...data, equipment: scopedEquipment, entries: data.entries.filter((en) => scopedIds.has(en.equipmentId)) };
  }, [data, scopedEquipment, scopedIds]);

  const latestEntryByEquipment = useMemo(() => {
    if (!data) return {};
    const map = {};
    for (const e of data.entries) {
      if (!map[e.equipmentId] || e.date > map[e.equipmentId].date) map[e.equipmentId] = e;
    }
    return map;
  }, [data]);

  const projectStats = useMemo(() => {
    if (!data) return [];
    return data.projects.map((p) => {
      const eqs = scopedEquipment.filter((e) => e.project === p);
      const good = eqs.filter((e) => e.status === STATUS.GOOD).length;
      const breakdown = eqs.filter((e) => e.status === STATUS.BREAKDOWN).length;
      const maintenance = eqs.filter((e) => e.status === STATUS.MAINTENANCE).length;
      const rent = eqs.reduce((s, e) => s + (Number(e.rate) || 0), 0);
      const tpiSoon = eqs.filter((e) => { const dd = daysFromToday(e.tpiExpiry); return dd !== null && dd >= 0 && dd <= 30; }).length;
      return { project: p, count: eqs.length, good, breakdown, maintenance, rent, tpiSoon };
    });
  }, [data, scopedEquipment]);

  const overall = useMemo(() => {
    if (!data) return null;
    const eqs = scopedEquipment;
    const good = eqs.filter((e) => e.status === STATUS.GOOD).length;
    const breakdown = eqs.filter((e) => e.status === STATUS.BREAKDOWN).length;
    const maintenance = eqs.filter((e) => e.status === STATUS.MAINTENANCE).length;
    const idle = eqs.filter((e) => e.status === STATUS.IDLE).length;
    const dehired = eqs.filter((e) => e.status === STATUS.DEHIRED).length;
    const rent = eqs.reduce((s, e) => s + (Number(e.rate) || 0), 0);
    const tpiSoon = eqs.filter((e) => { const dd = daysFromToday(e.tpiExpiry); return dd !== null && dd >= 0 && dd <= 30; }).length;
    const notUpdatedToday = eqs.filter((e) => { const le = latestEntryByEquipment[e.id]; return !le || le.date !== todayStr(); }).length;
    return { total: eqs.length, good, breakdown, maintenance, idle, dehired, rent, tpiSoon, notUpdatedToday };
  }, [data, scopedEquipment, latestEntryByEquipment]);

  if (!loaded || !data) {
    return (
      <div className="sc-root sc-loading">
        <Style />
        <div className="sc-loading-inner"><RefreshCw className="sc-spin" size={22} /><span>Loading equipment data…</span></div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className="sc-root">
        <Style />
        <AuthGate users={data.users} projects={data.projects} onRegister={registerUser} onLogin={setCurrentUser} />
      </div>
    );
  }

  const isAdmin = currentUser.role === ROLE_ADMIN;

  const NAV = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "daily", label: "Daily Update", icon: ClipboardCheck },
    { id: "projects", label: "By Project", icon: FolderKanban },
    { id: "equipment", label: "Equipment", icon: Truck },
    { id: "history", label: "History", icon: HistoryIcon },
    { id: "analysis", label: "Analysis", icon: BarChart3 },
    ...(isAdmin ? [{ id: "users", label: "Users", icon: UsersIcon }] : []),
  ];

  return (
    <div className="sc-root">
      <Style />
      <div className="sc-shell">
        <header className="sc-topbar">
          <div className="sc-topbar-left">
            <Logo size={32} />
            <div className="sc-topbar-brand-text">
              <div className="sc-brand-name">SISCOL</div>
              <div className="sc-brand-sub">P&amp;M Command Center</div>
            </div>
          </div>

          <nav className="sc-topnav">
            {NAV.map((n) => {
              const Icon = n.icon;
              return (
                <button key={n.id} className={classNames("sc-topnav-item", tab === n.id && "sc-topnav-item-active")} title={n.label} onClick={() => setTab(n.id)}>
                  <Icon size={16} strokeWidth={2} /><span>{n.label}</span>
                </button>
              );
            })}
          </nav>

          <div className="sc-topbar-right">
            <div className="sc-segmented sc-segmented-compact">
              <button className={classNames("sc-segmented-opt", scope === SCOPE_SISCOL && "sc-segmented-opt-active")} onClick={() => setScope(SCOPE_SISCOL)}>SISCOL</button>
              <button className={classNames("sc-segmented-opt", scope === SCOPE_ALL && "sc-segmented-opt-active")} onClick={() => setScope(SCOPE_ALL)}>All</button>
            </div>

            <button className="sc-export-btn" title="Export to Excel" onClick={() => { exportToExcel(scopedData, latestEntryByEquipment, scope); showToast("Excel file downloading…"); }}>
              <FileDown size={14} /> <span className="sc-hide-mid">Export</span>
            </button>

            <span
              className="sc-sync-dot"
              title={syncState === "saving" ? "Syncing…" : syncState === "error" ? "Sync failed — retrying" : "Live for all registered users"}
            >
              {syncState === "saving" ? <RefreshCw size={14} className="sc-spin" /> : syncState === "error" ? <CloudOff size={14} color="#E5544B" /> : <Cloud size={14} color="#3FBE83" />}
            </span>

            <div className="sc-user-chip" title={currentUser.name}>
              <span className="sc-user-avatar"><UserCircle2 size={16} /></span>
              <div className="sc-user-chip-text sc-hide-mid">
                <div className="sc-strong" style={{ fontSize: 12.5 }}>{currentUser.name}</div>
                <div className="sc-muted" style={{ fontSize: 11 }}>{isAdmin ? "Admin" : `Incharge · ${currentUser.project}`}</div>
              </div>
              <button className="sc-icon-btn" title="Sign out" onClick={() => setCurrentUser(null)}><LogOut size={15} /></button>
            </div>
          </div>
        </header>

        <main className="sc-main">
          {tab === "dashboard" && (
            <Dashboard
              data={scopedData}
              scope={scope}
              overall={overall}
              projectStats={projectStats}
              latestEntryByEquipment={latestEntryByEquipment}
              goToDaily={() => setTab("daily")}
              onDrill={setDetailFilter}
            />
          )}
          {tab === "daily" && (
            <DailyUpdate data={scopedData} currentUser={currentUser} latestEntryByEquipment={latestEntryByEquipment} onLog={logStatus} showToast={showToast} />
          )}
          {tab === "projects" && (
            <ProjectsTab data={data} scope={scope} isAdmin={isAdmin} projectStats={projectStats} onAdd={addProject} onDelete={deleteProject} showToast={showToast} onOpenProject={setOpenProject} />
          )}
          {tab === "equipment" && (
            <EquipmentTab data={data} scope={scope} isAdmin={isAdmin} rentedByOptions={rentedByOptions} onAdd={addEquipment} onUpdate={updateEquipment} onDelete={deleteEquipment} onAddProject={addProject} showToast={showToast} />
          )}
          {tab === "history" && <HistoryTab data={data} scope={scope} rentedByOptions={rentedByOptions} />}
          {tab === "analysis" && <AnalysisTab data={scopedData} scope={scope} projectStats={projectStats} overall={overall} />}
          {tab === "users" && isAdmin && <UsersTab data={data} setData={setData} currentUser={currentUser} showToast={showToast} />}
        </main>
      </div>

      {detailFilter && (
        <EquipmentDetailModal
          status={detailFilter}
          equipment={scopedData.equipment.filter((e) => e.status === detailFilter)}
          latestEntryByEquipment={latestEntryByEquipment}
          onClose={() => setDetailFilter(null)}
        />
      )}
      {openProject && (
        <ProjectDetailModal
          project={openProject}
          data={data}
          scope={scope}
          currentUser={currentUser}
          latestEntryByEquipment={latestEntryByEquipment}
          onLog={logStatus}
          onClose={() => setOpenProject(null)}
          showToast={showToast}
        />
      )}
      {toast && <div className="sc-toast">{toast}</div>}
    </div>
  );
}

/* ==================================================================== */
/* Dashboard                                                            */
/* ==================================================================== */

function Dashboard({ data, scope, overall, projectStats, latestEntryByEquipment, goToDaily, onDrill }) {
  const pending = data.equipment.filter((e) => {
    const le = latestEntryByEquipment[e.id];
    return (!le || le.date !== todayStr()) && e.status !== STATUS.DEHIRED;
  });

  const tpiList = data.equipment
    .map((e) => ({ e, dd: daysFromToday(e.tpiExpiry) }))
    .filter((x) => x.dd !== null && x.dd <= 30)
    .sort((a, b) => a.dd - b.dd)
    .slice(0, 8);

  return (
    <div className="sc-page">
      <div className="sc-hero">
        <div className="sc-hero-glow" />
        <div className="sc-hero-inner">
          <div className="sc-hero-brand">
            <Logo size={40} />
            <div>
              <div className="sc-hero-title">SISCOL Plant &amp; Machinery</div>
              <div className="sc-hero-sub">
                {scope === SCOPE_SISCOL ? "SISCOL-rented equipment only" : "All contractors"} · {overall.total} equipment · {data.projects.length} project sites
              </div>
            </div>
          </div>
          <div className="sc-hero-date">{new Date().toLocaleDateString("en-IN", { weekday: "long", day: "2-digit", month: "long", year: "numeric" })}</div>
        </div>
      </div>

      <div className="sc-stat-grid">
        <StatCard label="Total equipment" value={overall.total} icon={Truck} />
        <StatCard label="Operational" value={overall.good} accent="#3FBE83" icon={Check} onClick={() => onDrill(STATUS.GOOD)} />
        <StatCard label="Breakdown" value={overall.breakdown} accent="#E5544B" icon={AlertTriangle} sub={overall.breakdown > 0 ? "tap to see which" : undefined} onClick={() => onDrill(STATUS.BREAKDOWN)} />
        <StatCard label="Under maintenance" value={overall.maintenance} accent="#E3A438" icon={Wrench} sub={overall.maintenance > 0 ? "tap to see which" : undefined} onClick={() => onDrill(STATUS.MAINTENANCE)} />
        <StatCard label="Total rent (INR)" value={fmtINR(overall.rent)} icon={IndianRupee} />
        <StatCard label="TPI expiring ≤30 days" value={overall.tpiSoon} accent={overall.tpiSoon > 0 ? "#E3A438" : undefined} icon={AlertTriangle} />
      </div>

      <div className="sc-grid-2">
        <div className="sc-panel">
          <div className="sc-panel-head"><h3>Project-wise summary</h3></div>
          <div className="sc-table-wrap">
            <table className="sc-table">
              <thead>
                <tr>
                  <th>Project</th><th className="sc-num">Equipment</th><th className="sc-num">Rent (INR)</th>
                  <th className="sc-num">Operational</th><th className="sc-num">Breakdown</th><th className="sc-num">TPI ≤30d</th>
                </tr>
              </thead>
              <tbody>
                {projectStats.map((p) => (
                  <tr key={p.project}>
                    <td className="sc-strong">{p.project}</td>
                    <td className="sc-num sc-mono">{p.count}</td>
                    <td className="sc-num sc-mono">{fmtINR(p.rent)}</td>
                    <td className="sc-num sc-mono" style={{ color: "#3FBE83" }}>{p.good}</td>
                    <td className="sc-num sc-mono" style={{ color: p.breakdown ? "#E5544B" : undefined }}>{p.breakdown}</td>
                    <td className="sc-num sc-mono" style={{ color: p.tpiSoon ? "#E3A438" : undefined }}>{p.tpiSoon}</td>
                  </tr>
                ))}
                {projectStats.length === 0 && <tr><td colSpan={6} className="sc-empty">No projects yet.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>

        <div className="sc-panel">
          <div className="sc-panel-head"><h3>TPI expiry watchlist</h3></div>
          <div className="sc-list">
            {tpiList.length === 0 && <div className="sc-empty">Nothing expiring within 30 days.</div>}
            {tpiList.map(({ e, dd }) => (
              <div key={e.id} className="sc-list-row">
                <div><div className="sc-strong">{e.type}</div><div className="sc-muted sc-mono" style={{ fontSize: 12 }}>{e.vehicleNo || e.id} · {e.project}</div></div>
                <div className={classNames("sc-tpi-chip", dd < 0 && "sc-tpi-chip-overdue")}>{dd < 0 ? `${Math.abs(dd)}d overdue` : dd === 0 ? "today" : `${dd}d left`}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="sc-panel">
        <div className="sc-panel-head">
          <h3>Pending today's update</h3>
          <span className="sc-muted">{pending.length} of {overall.total} equipment not yet logged today</span>
        </div>
        {pending.length === 0 ? (
          <div className="sc-empty">All equipment has been updated for today. Nice work.</div>
        ) : (
          <>
            <div className="sc-chip-row">
              {pending.slice(0, 14).map((e) => <span key={e.id} className="sc-chip">{e.type} · {e.project}</span>)}
              {pending.length > 14 && <span className="sc-chip sc-chip-muted">+{pending.length - 14} more</span>}
            </div>
            <button className="sc-btn sc-btn-primary" style={{ marginTop: 14 }} onClick={goToDaily}>Go to Daily Update</button>
          </>
        )}
      </div>
    </div>
  );
}

/* ==================================================================== */
/* Equipment detail drill-down (from clickable stat cards)              */
/* ==================================================================== */

function EquipmentDetailModal({ status, equipment, latestEntryByEquipment, onClose }) {
  const st = STATUS_STYLE[status];
  return (
    <Modal title={`${st.label} equipment`} subtitle={`${equipment.length} equipment currently marked ${st.label.toLowerCase()}`} onClose={onClose} wide>
      {equipment.length === 0 ? (
        <div className="sc-empty">Nothing here right now.</div>
      ) : (
        <div className="sc-detail-list">
          {equipment.map((e) => {
            const le = latestEntryByEquipment[e.id];
            return (
              <div key={e.id} className="sc-detail-card">
                <div className="sc-detail-card-head">
                  <div>
                    <div className="sc-strong">{e.type}</div>
                    <div className="sc-muted sc-mono" style={{ fontSize: 12 }}>{e.id} · {e.project}</div>
                  </div>
                  <StatusBadge status={e.status} size="sm" />
                </div>
                <div className="sc-detail-grid">
                  <div><span className="sc-muted">Vendor</span><span>{e.vendor}</span></div>
                  <div><span className="sc-muted">Rented by</span><span>{e.rentedBy || "SISCOL"}</span></div>
                  <div><span className="sc-muted">Vehicle no.</span><span className="sc-mono">{e.vehicleNo || "—"}</span></div>
                  <div><span className="sc-muted">Contact</span><span className="sc-mono">{e.contact || "—"}</span></div>
                  <div><span className="sc-muted">Rate (INR)</span><span className="sc-mono">{fmtINR(e.rate)}</span></div>
                  <div><span className="sc-muted">Since</span><span>{le ? fmtDateShort(le.date) : "—"}</span></div>
                </div>
                {le && (le.reason || le.remark) && (
                  <div className="sc-detail-note">
                    {le.reason && <div><span className="sc-muted">Reason: </span>{le.reason}</div>}
                    {le.remark && <div><span className="sc-muted">Remark: </span>{le.remark}</div>}
                    <div className="sc-muted" style={{ fontSize: 11.5, marginTop: 4 }}>Logged by {le.loggedBy || "unknown"} · {fmtDateTime(le.loggedAt)}</div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </Modal>
  );
}

/* ==================================================================== */
/* Daily Update                                                         */
/* ==================================================================== */

function DailyUpdate({ data, currentUser, latestEntryByEquipment, onLog, showToast }) {
  const isAdmin = currentUser.role === ROLE_ADMIN;
  const [projectFilter, setProjectFilter] = useState(isAdmin ? "All" : currentUser.project);
  const [statusFilter, setStatusFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [openId, setOpenId] = useState(null);

  const rows = data.equipment.filter((e) => {
    if (projectFilter !== "All" && e.project !== projectFilter) return false;
    if (statusFilter !== "All" && e.status !== statusFilter) return false;
    if (search.trim()) {
      const q = search.toLowerCase();
      if (!`${e.type} ${e.vendor} ${e.vehicleNo} ${e.id}`.toLowerCase().includes(q)) return false;
    }
    return true;
  });

  return (
    <div className="sc-page">
      <PageHead
        title="Daily Update"
        sub={isAdmin ? "Log today's (or a past) status for any equipment." : `Logging status for ${currentUser.project} — every save is added to that equipment's history.`}
      />

      <div className="sc-toolbar">
        <div className="sc-search"><Search size={15} /><input placeholder="Search equipment, vendor, vehicle no…" value={search} onChange={(e) => setSearch(e.target.value)} /></div>
        <Select value={projectFilter} onChange={setProjectFilter} options={isAdmin ? ["All", ...data.projects] : [currentUser.project]} disabled={!isAdmin} />
        <Select value={statusFilter} onChange={setStatusFilter} options={["All", ...STATUS_ORDER]} />
      </div>

      <div className="sc-cardlist">
        {rows.map((e) => {
          const le = latestEntryByEquipment[e.id];
          const updatedToday = le && le.date === todayStr();
          return (
            <div key={e.id} className={classNames("sc-eqcard", !updatedToday && "sc-eqcard-pending")}>
              <div className="sc-eqcard-main">
                <div className="sc-eqcard-title"><span className="sc-strong">{e.type}</span><span className="sc-muted sc-mono" style={{ fontSize: 12 }}>{e.id}</span></div>
                <div className="sc-eqcard-meta"><span>{e.project}</span><span className="sc-dot">·</span><span className="sc-mono">{e.vehicleNo || "—"}</span><span className="sc-dot">·</span><span>{e.vendor}</span></div>
                <div className="sc-eqcard-meta sc-muted" style={{ fontSize: 12 }}>Last logged: {le ? `${fmtDateShort(le.date)} — ${le.status} by ${le.loggedBy || "?"}` : "never"}</div>
              </div>
              <div className="sc-eqcard-side">
                <StatusBadge status={e.status} />
                <button className={classNames("sc-btn", updatedToday ? "sc-btn-ghost" : "sc-btn-primary")} onClick={() => setOpenId(e.id)}>{updatedToday ? "Update again" : "Log status"}</button>
              </div>
            </div>
          );
        })}
        {rows.length === 0 && <div className="sc-empty" style={{ padding: 32 }}>No equipment matches these filters.</div>}
      </div>

      {openId && (
        <LogStatusModal
          equipment={data.equipment.find((e) => e.id === openId)}
          onClose={() => setOpenId(null)}
          onSave={(entry) => { onLog(openId, entry, currentUser); showToast("Status logged"); setOpenId(null); }}
        />
      )}
    </div>
  );
}

function LogStatusModal({ equipment, onClose, onSave }) {
  const [date, setDate] = useState(todayStr());
  const [status, setStatus] = useState(equipment.status || STATUS.GOOD);
  const [reason, setReason] = useState("");
  const [remark, setRemark] = useState("");

  return (
    <Modal title={`Log status — ${equipment.type} (${equipment.id})`} onClose={onClose}>
      <div className="sc-form">
        <Field label="Date"><input type="date" className="sc-input" value={date} max={todayStr()} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="Status" group>
          <div className="sc-status-picker">
            {STATUS_ORDER.map((s) => {
              const st = STATUS_STYLE[s]; const Icon = st.icon;
              return (
                <button type="button" key={s} className={classNames("sc-status-opt", status === s && "sc-status-opt-active")} style={status === s ? { borderColor: st.color, color: st.color, background: st.bg } : undefined} onClick={() => setStatus(s)}>
                  <Icon size={13} />{st.label}
                </button>
              );
            })}
          </div>
        </Field>
        {(status === STATUS.BREAKDOWN || status === STATUS.MAINTENANCE) && (
          <Field label="Reason"><input className="sc-input" placeholder="e.g. hydraulic hose leak, engine overheating…" value={reason} onChange={(e) => setReason(e.target.value)} /></Field>
        )}
        <Field label="Remark" hint="Optional notes visible in the equipment's history"><textarea className="sc-textarea" rows={3} value={remark} onChange={(e) => setRemark(e.target.value)} placeholder="Any additional detail…" /></Field>
        <div className="sc-form-actions">
          <button className="sc-btn sc-btn-ghost" onClick={onClose}>Cancel</button>
          <button className="sc-btn sc-btn-primary" onClick={() => onSave({ date, status, reason, remark })}>Save entry</button>
        </div>
      </div>
    </Modal>
  );
}

/* ==================================================================== */
/* Equipment tab                                                        */
/* ==================================================================== */

function EquipmentTab({ data, scope, isAdmin, rentedByOptions, onAdd, onUpdate, onDelete, onAddProject, showToast }) {
  const [projectFilter, setProjectFilter] = useState("All");
  const [rentedByFilter, setRentedByFilter] = useState(scope === SCOPE_SISCOL ? "SISCOL" : "All");
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [editId, setEditId] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const rows = data.equipment.filter((e) => {
    if (projectFilter !== "All" && e.project !== projectFilter) return false;
    if (rentedByFilter !== "All" && (e.rentedBy || "SISCOL") !== rentedByFilter) return false;
    if (search.trim()) {
      const q = search.toLowerCase();
      if (!`${e.type} ${e.vendor} ${e.vehicleNo} ${e.id} ${e.model}`.toLowerCase().includes(q)) return false;
    }
    return true;
  });

  const siscolCount = data.equipment.filter((e) => (e.rentedBy || "SISCOL") === "SISCOL").length;

  return (
    <div className="sc-page">
      <PageHead
        title="Equipment"
        sub={`${data.equipment.length} equipment across ${data.projects.length} projects · ${siscolCount} rented by SISCOL, ${data.equipment.length - siscolCount} by other contractors`}
        action={isAdmin && <button className="sc-btn sc-btn-primary" onClick={() => setShowAdd(true)}><Plus size={15} /> Add equipment</button>}
      />

      <div className="sc-toolbar">
        <div className="sc-search"><Search size={15} /><input placeholder="Search type, vendor, vehicle no, model…" value={search} onChange={(e) => setSearch(e.target.value)} /></div>
        <Select value={projectFilter} onChange={setProjectFilter} options={["All", ...data.projects]} />
        <Select value={rentedByFilter} onChange={setRentedByFilter} options={["All", ...rentedByOptions]} />
      </div>

      <div className="sc-panel">
        <div className="sc-table-wrap">
          <table className="sc-table">
            <thead>
              <tr>
                <th>ID</th><th>Project</th><th>Type</th><th>Vendor</th><th>Rented by</th>
                <th>Vehicle no.</th><th className="sc-num">Rate (INR)</th><th>Status</th><th>TPI expiry</th>{isAdmin && <th></th>}
              </tr>
            </thead>
            <tbody>
              {rows.map((e) => (
                <tr key={e.id}>
                  <td className="sc-mono sc-muted">{e.id}</td>
                  <td>{e.project}</td>
                  <td className="sc-strong">{e.type}</td>
                  <td>{e.vendor}</td>
                  <td>{(e.rentedBy || "SISCOL") === "SISCOL" ? <span className="sc-rentedby-siscol">SISCOL</span> : <span className="sc-rentedby-other">{e.rentedBy}</span>}</td>
                  <td className="sc-mono">{e.vehicleNo || "—"}</td>
                  <td className="sc-num sc-mono">{fmtINR(e.rate)}</td>
                  <td><StatusBadge status={e.status} size="sm" /></td>
                  <td className="sc-mono" style={{ fontSize: 12.5 }}>{e.tpiExpiry ? fmtDateShort(e.tpiExpiry) : "—"}</td>
                  {isAdmin && (
                    <td>
                      <div className="sc-row-actions">
                        <button className="sc-icon-btn" onClick={() => setEditId(e.id)}><Pencil size={14} /></button>
                        <button className="sc-icon-btn sc-icon-btn-danger" onClick={() => setConfirmDelete(e.id)}><Trash2 size={14} /></button>
                      </div>
                    </td>
                  )}
                </tr>
              ))}
              {rows.length === 0 && <tr><td colSpan={isAdmin ? 10 : 9} className="sc-empty">No equipment matches these filters.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {showAdd && (
        <EquipmentForm projects={data.projects} rentedByOptions={rentedByOptions} onClose={() => setShowAdd(false)} onAddProject={onAddProject} onSave={(eq) => { onAdd(eq); showToast("Equipment added"); setShowAdd(false); }} />
      )}
      {editId && (
        <EquipmentForm projects={data.projects} rentedByOptions={rentedByOptions} initial={data.equipment.find((e) => e.id === editId)} onClose={() => setEditId(null)} onAddProject={onAddProject} onSave={(eq) => { onUpdate(editId, eq); showToast("Equipment updated"); setEditId(null); }} />
      )}
      {confirmDelete && (
        <Modal title="Remove equipment" onClose={() => setConfirmDelete(null)}>
          <p style={{ color: "#C7CCD3", lineHeight: 1.5 }}>This removes the equipment and its full status history. This can't be undone.</p>
          <div className="sc-form-actions">
            <button className="sc-btn sc-btn-ghost" onClick={() => setConfirmDelete(null)}>Cancel</button>
            <button className="sc-btn sc-btn-danger" onClick={() => { onDelete(confirmDelete); setConfirmDelete(null); showToast("Equipment removed"); }}>Remove</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

function EquipmentForm({ projects, rentedByOptions, initial, onClose, onSave, onAddProject }) {
  const [project, setProject] = useState(initial?.project || projects[0] || "");
  const [newProject, setNewProject] = useState("");
  const [type, setType] = useState(initial?.type || EQUIPMENT_TYPES[0]);
  const [customType, setCustomType] = useState(initial && !EQUIPMENT_TYPES.includes(initial.type) ? initial.type : "");
  const [vendor, setVendor] = useState(initial?.vendor || "");
  const [rentedBy, setRentedBy] = useState(initial?.rentedBy || "SISCOL");
  const [newRentedBy, setNewRentedBy] = useState("");
  const [rate, setRate] = useState(initial?.rate ?? "");
  const [shifts, setShifts] = useState(initial?.shifts ?? 1);
  const [model, setModel] = useState(initial?.model || "");
  const [yom, setYom] = useState(initial?.yom || "");
  const [vehicleNo, setVehicleNo] = useState(initial?.vehicleNo || "");
  const [contact, setContact] = useState(initial?.contact || "");
  const [tpiExpiry, setTpiExpiry] = useState(initial?.tpiExpiry || "");

  function submit() {
    let finalProject = project;
    if (project === "__new__") {
      const clean = newProject.trim();
      if (!clean) return;
      onAddProject(clean);
      finalProject = clean;
    }
    const finalType = type === "Other" ? (customType.trim() || "Other") : type;
    let finalRentedBy = rentedBy;
    if (rentedBy === "__new__") {
      finalRentedBy = newRentedBy.trim();
      if (!finalRentedBy) return;
    }
    if (!finalProject || !finalType || !vendor.trim()) return;
    onSave({
      project: finalProject,
      type: finalType,
      vendor: vendor.trim(),
      rentedBy: finalRentedBy,
      rate: Number(rate) || 0,
      shifts: Number(shifts) || 1,
      model: model.trim(),
      yom: yom ? (isNaN(Number(yom)) ? yom : Number(yom)) : null,
      vehicleNo: vehicleNo.trim(),
      contact: contact.trim(),
      tpiExpiry: tpiExpiry || null,
    });
  }

  return (
    <Modal title={initial ? "Edit equipment" : "Add equipment"} onClose={onClose} wide>
      <div className="sc-form sc-form-grid">
        <Field label="Project">
          <div className="sc-select" style={{ width: "100%" }}>
            <select value={project} onChange={(e) => setProject(e.target.value)}>
              {projects.map((p) => <option key={p} value={p}>{p}</option>)}
              <option value="__new__">+ New project…</option>
            </select>
            <ChevronDown size={14} />
          </div>
          {project === "__new__" && <input className="sc-input" style={{ marginTop: 8 }} placeholder="New project name" value={newProject} onChange={(e) => setNewProject(e.target.value)} />}
        </Field>
        <Field label="Equipment type">
          <div className="sc-select" style={{ width: "100%" }}>
            <select value={type} onChange={(e) => setType(e.target.value)}>{EQUIPMENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select>
            <ChevronDown size={14} />
          </div>
          {type === "Other" && <input className="sc-input" style={{ marginTop: 8 }} placeholder="Specify type" value={customType} onChange={(e) => setCustomType(e.target.value)} />}
        </Field>
        <Field label="Vendor / owner"><input className="sc-input" value={vendor} onChange={(e) => setVendor(e.target.value)} /></Field>
        <Field label="Rented by" hint="Who this equipment is rented for — SISCOL, or another contractor on site">
          <div className="sc-select" style={{ width: "100%" }}>
            <select value={rentedBy} onChange={(e) => setRentedBy(e.target.value)}>
              {rentedByOptions.map((r) => <option key={r} value={r}>{r}</option>)}
              <option value="__new__">+ Other contractor…</option>
            </select>
            <ChevronDown size={14} />
          </div>
          {rentedBy === "__new__" && <input className="sc-input" style={{ marginTop: 8 }} placeholder="Contractor name" value={newRentedBy} onChange={(e) => setNewRentedBy(e.target.value)} />}
        </Field>
        <Field label="Rate (INR)"><input className="sc-input" type="number" value={rate} onChange={(e) => setRate(e.target.value)} /></Field>
        <Field label="Shifts / hrs"><input className="sc-input" type="number" value={shifts} onChange={(e) => setShifts(e.target.value)} /></Field>
        <Field label="Model"><input className="sc-input" value={model} onChange={(e) => setModel(e.target.value)} /></Field>
        <Field label="Year of manufacture"><input className="sc-input" value={yom} onChange={(e) => setYom(e.target.value)} /></Field>
        <Field label="Vehicle / serial no."><input className="sc-input" value={vehicleNo} onChange={(e) => setVehicleNo(e.target.value)} /></Field>
        <Field label="Contact number"><input className="sc-input" value={contact} onChange={(e) => setContact(e.target.value)} /></Field>
        <Field label="TPI expiry" hint="Leave blank if not applicable"><input className="sc-input" type="date" value={tpiExpiry || ""} onChange={(e) => setTpiExpiry(e.target.value)} /></Field>
      </div>
      <div className="sc-form-actions">
        <button className="sc-btn sc-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="sc-btn sc-btn-primary" onClick={submit}>{initial ? "Save changes" : "Add equipment"}</button>
      </div>
    </Modal>
  );
}

/* ==================================================================== */
/* Projects tab — project-wise drill down                               */
/* ==================================================================== */

function ProjectsTab({ data, scope, isAdmin, projectStats, onAdd, onDelete, showToast, onOpenProject }) {
  const [showAdd, setShowAdd] = useState(false);
  const [name, setName] = useState("");
  const [error, setError] = useState("");

  function submit() {
    if (!name.trim()) return;
    const ok = onAdd(name);
    if (!ok) { setError("A project with this name already exists."); return; }
    showToast("Project added"); setName(""); setError(""); setShowAdd(false);
  }

  return (
    <div className="sc-page">
      <PageHead
        title="By Project"
        sub={`${data.projects.length} active project sites · figures reflect ${scope === SCOPE_SISCOL ? "SISCOL-rented equipment only" : "all contractors"} · tap a card to see its full equipment list`}
        action={isAdmin && <button className="sc-btn sc-btn-primary" onClick={() => setShowAdd(true)}><Plus size={15} /> Add project</button>}
      />

      <div className="sc-project-grid">
        {projectStats.map((p) => (
          <button key={p.project} className="sc-project-card sc-project-card-clickable" onClick={() => onOpenProject(p.project)}>
            <div className="sc-project-card-head">
              <h3>{p.project}</h3>
              {isAdmin && (
                <span
                  className="sc-icon-btn sc-icon-btn-danger"
                  title="Delete project"
                  onClick={(ev) => {
                    ev.stopPropagation();
                    const ok = onDelete(p.project);
                    if (!ok) showToast("Move or remove its equipment before deleting a project");
                    else showToast("Project removed");
                  }}
                >
                  <Trash2 size={14} />
                </span>
              )}
            </div>
            <div className="sc-project-stats">
              <div><span className="sc-muted">Equipment</span><span className="sc-strong sc-mono">{p.count}</span></div>
              <div><span className="sc-muted">Rent (INR)</span><span className="sc-strong sc-mono">{fmtINR(p.rent)}</span></div>
              <div><span className="sc-muted">Operational</span><span className="sc-mono" style={{ color: "#3FBE83" }}>{p.good}</span></div>
              <div><span className="sc-muted">Breakdown</span><span className="sc-mono" style={{ color: p.breakdown ? "#E5544B" : undefined }}>{p.breakdown}</span></div>
            </div>
            <div className="sc-project-card-cta">Open equipment list <ChevronRight size={13} /></div>
          </button>
        ))}
        {projectStats.length === 0 && <div className="sc-empty">No projects yet — add one to get started.</div>}
      </div>

      {showAdd && (
        <Modal title="Add project" onClose={() => setShowAdd(false)}>
          <div className="sc-form">
            <Field label="Project name"><input className="sc-input" autoFocus value={name} onChange={(e) => { setName(e.target.value); setError(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} /></Field>
            {error && <div className="sc-error">{error}</div>}
            <div className="sc-form-actions"><button className="sc-btn sc-btn-ghost" onClick={() => setShowAdd(false)}>Cancel</button><button className="sc-btn sc-btn-primary" onClick={submit}>Add project</button></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function ProjectDetailModal({ project, data, scope, currentUser, latestEntryByEquipment, onLog, onClose, showToast }) {
  const [openId, setOpenId] = useState(null);
  const eqs = data.equipment.filter((e) => e.project === project && (scope === SCOPE_ALL || (e.rentedBy || "SISCOL") === "SISCOL"));
  const canLog = currentUser.role === ROLE_ADMIN || currentUser.project === project;

  return (
    <Modal title={project} subtitle={`${eqs.length} equipment on this site`} onClose={onClose} wide>
      <div className="sc-detail-list">
        {eqs.map((e) => {
          const le = latestEntryByEquipment[e.id];
          return (
            <div key={e.id} className="sc-detail-card">
              <div className="sc-detail-card-head">
                <div><div className="sc-strong">{e.type}</div><div className="sc-muted sc-mono" style={{ fontSize: 12 }}>{e.id} · {e.vehicleNo || "no reg. no."}</div></div>
                <StatusBadge status={e.status} size="sm" />
              </div>
              <div className="sc-detail-grid">
                <div><span className="sc-muted">Vendor</span><span>{e.vendor}</span></div>
                <div><span className="sc-muted">Rented by</span><span>{e.rentedBy || "SISCOL"}</span></div>
                <div><span className="sc-muted">Rate (INR)</span><span className="sc-mono">{fmtINR(e.rate)}</span></div>
                <div><span className="sc-muted">Last logged</span><span>{le ? `${fmtDateShort(le.date)} by ${le.loggedBy || "?"}` : "never"}</span></div>
              </div>
              {canLog && <button className="sc-btn sc-btn-ghost" style={{ marginTop: 10 }} onClick={() => setOpenId(e.id)}>Log status</button>}
            </div>
          );
        })}
        {eqs.length === 0 && <div className="sc-empty">No equipment recorded on this project yet.</div>}
      </div>
      {openId && (
        <LogStatusModal
          equipment={eqs.find((e) => e.id === openId)}
          onClose={() => setOpenId(null)}
          onSave={(entry) => { onLog(openId, entry, currentUser); showToast("Status logged"); setOpenId(null); }}
        />
      )}
    </Modal>
  );
}

/* ==================================================================== */
/* Users tab (admin only)                                               */
/* ==================================================================== */

function UsersTab({ data, setData, currentUser, showToast }) {
  const [confirmRemove, setConfirmRemove] = useState(null);

  function removeUser(id) {
    setData((d) => ({ ...d, users: d.users.filter((u) => u.id !== id) }));
    showToast("User removed");
    setConfirmRemove(null);
  }

  return (
    <div className="sc-page">
      <PageHead title="Users" sub={`${data.users.length} registered users · admins see every site, incharges are limited to their own`} />
      <div className="sc-panel">
        <div className="sc-table-wrap">
          <table className="sc-table">
            <thead><tr><th>Name</th><th>Role</th><th>Site</th><th>Registered</th><th></th></tr></thead>
            <tbody>
              {data.users.map((u) => (
                <tr key={u.id}>
                  <td className="sc-strong">{u.name}{u.id === currentUser.id && <span className="sc-muted"> (you)</span>}</td>
                  <td>{u.role === ROLE_ADMIN ? <span className="sc-rentedby-siscol"><ShieldCheck size={12} style={{ verticalAlign: "-2px", marginRight: 3 }} />Admin</span> : "Site incharge"}</td>
                  <td>{u.project || "All sites"}</td>
                  <td className="sc-mono" style={{ fontSize: 12 }}>{fmtDateShort(u.createdAt?.slice(0, 10))}</td>
                  <td>{u.id !== currentUser.id && <button className="sc-icon-btn sc-icon-btn-danger" onClick={() => setConfirmRemove(u.id)}><Trash2 size={14} /></button>}</td>
                </tr>
              ))}
              {data.users.length === 0 && <tr><td colSpan={5} className="sc-empty">No one registered yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
      {confirmRemove && (
        <Modal title="Remove user" onClose={() => setConfirmRemove(null)}>
          <p style={{ color: "#C7CCD3", lineHeight: 1.5 }}>This revokes their access. Their past status entries stay in history under their name.</p>
          <div className="sc-form-actions">
            <button className="sc-btn sc-btn-ghost" onClick={() => setConfirmRemove(null)}>Cancel</button>
            <button className="sc-btn sc-btn-danger" onClick={() => removeUser(confirmRemove)}>Remove</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ==================================================================== */
/* History tab                                                          */
/* ==================================================================== */

function HistoryTab({ data, scope, rentedByOptions }) {
  const [projectFilter, setProjectFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [rentedByFilter, setRentedByFilter] = useState(scope === SCOPE_SISCOL ? "SISCOL" : "All");
  const [search, setSearch] = useState("");

  const eqById = useMemo(() => { const m = {}; data.equipment.forEach((e) => (m[e.id] = e)); return m; }, [data.equipment]);

  const rows = data.entries
    .map((en) => ({ ...en, eq: eqById[en.equipmentId] }))
    .filter((r) => r.eq)
    .filter((r) => (projectFilter === "All" ? true : r.eq.project === projectFilter))
    .filter((r) => (statusFilter === "All" ? true : r.status === statusFilter))
    .filter((r) => (rentedByFilter === "All" ? true : (r.eq.rentedBy || "SISCOL") === rentedByFilter))
    .filter((r) => {
      if (!search.trim()) return true;
      const q = search.toLowerCase();
      return `${r.eq.type} ${r.eq.vendor} ${r.eq.vehicleNo} ${r.reason} ${r.remark} ${r.loggedBy || ""}`.toLowerCase().includes(q);
    })
    .sort((a, b) => (a.date === b.date ? (b.loggedAt || "").localeCompare(a.loggedAt || "") : b.date.localeCompare(a.date)));

  return (
    <div className="sc-page">
      <PageHead title="History" sub={`${data.entries.length} status entries logged in total`} />
      <div className="sc-toolbar">
        <div className="sc-search"><Search size={15} /><input placeholder="Search reason, remark, vendor, logged by…" value={search} onChange={(e) => setSearch(e.target.value)} /></div>
        <Select value={projectFilter} onChange={setProjectFilter} options={["All", ...data.projects]} />
        <Select value={statusFilter} onChange={setStatusFilter} options={["All", ...STATUS_ORDER]} />
        <Select value={rentedByFilter} onChange={setRentedByFilter} options={["All", ...rentedByOptions]} />
      </div>
      <div className="sc-panel">
        <div className="sc-table-wrap">
          <table className="sc-table">
            <thead><tr><th>Date</th><th>Project</th><th>Equipment</th><th>Rented by</th><th>Vehicle no.</th><th>Status</th><th>Reason</th><th>Remark</th><th>Logged by</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.id}>
                  <td className="sc-mono" style={{ fontSize: 12.5 }}>{fmtDateShort(r.date)}</td>
                  <td>{r.eq.project}</td>
                  <td className="sc-strong">{r.eq.type} <span className="sc-muted sc-mono" style={{ fontSize: 11.5 }}>{r.eq.id}</span></td>
                  <td>{(r.eq.rentedBy || "SISCOL") === "SISCOL" ? <span className="sc-rentedby-siscol">SISCOL</span> : <span className="sc-rentedby-other">{r.eq.rentedBy}</span>}</td>
                  <td className="sc-mono">{r.eq.vehicleNo || "—"}</td>
                  <td><StatusBadge status={r.status} size="sm" /></td>
                  <td className="sc-muted">{r.reason || "—"}</td>
                  <td className="sc-muted">{r.remark || "—"}</td>
                  <td className="sc-muted" style={{ fontSize: 12.5 }}>{r.loggedBy || "—"}</td>
                </tr>
              ))}
              {rows.length === 0 && <tr><td colSpan={9} className="sc-empty">No entries logged yet. Head to Daily Update to log the first one.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ==================================================================== */
/* Analysis tab                                                         */
/* ==================================================================== */

function AnalysisTab({ data, scope, projectStats, overall }) {
  const rentByProject = projectStats.map((p) => ({ name: p.project, rent: p.rent }));
  const countByProject = projectStats.map((p) => ({ name: p.project, equipment: p.count }));

  const last14 = useMemo(() => {
    const days = [];
    for (let i = 13; i >= 0; i--) { const d = new Date(); d.setDate(d.getDate() - i); days.push(localDateStr(d)); }
    return days.map((d) => {
      const entriesOnDay = data.entries.filter((e) => e.date === d);
      return { date: d.slice(5), breakdown: entriesOnDay.filter((e) => e.status === STATUS.BREAKDOWN).length, logged: entriesOnDay.length };
    });
  }, [data.entries]);

  const downtimeLeaders = useMemo(() => {
    const counts = {};
    data.entries.forEach((e) => { if (e.status === STATUS.BREAKDOWN) counts[e.equipmentId] = (counts[e.equipmentId] || 0) + 1; });
    return Object.entries(counts).map(([id, n]) => ({ eq: data.equipment.find((e) => e.id === id), n })).filter((x) => x.eq).sort((a, b) => b.n - a.n).slice(0, 8);
  }, [data.entries, data.equipment]);

  const typeBreakdown = useMemo(() => {
    const m = {};
    data.equipment.forEach((e) => { m[e.type] = (m[e.type] || 0) + 1; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([type, count]) => ({ type, count }));
  }, [data.equipment]);

  return (
    <div className="sc-page">
      <PageHead title="Analysis" sub={`Fleet-wide trends across every project · ${scope === SCOPE_SISCOL ? "SISCOL-rented equipment only" : "all contractors"}`} />

      <div className="sc-grid-2">
        <div className="sc-panel">
          <div className="sc-panel-head"><h3>Rent by project (INR)</h3></div>
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <BarChart data={rentByProject} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A2F37" vertical={false} />
                <XAxis dataKey="name" stroke="#8B93A1" fontSize={12} tickLine={false} axisLine={{ stroke: "#333941" }} />
                <YAxis stroke="#8B93A1" fontSize={11} tickLine={false} axisLine={false} tickFormatter={(v) => `${(v / 100000).toFixed(0)}L`} />
                <Tooltip contentStyle={{ background: "#1D2127", border: "1px solid #333941", borderRadius: 8, fontSize: 12.5 }} formatter={(v) => fmtINR(v)} />
                <Bar dataKey="rent" fill="#4A90D9" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="sc-panel">
          <div className="sc-panel-head"><h3>Equipment count by project</h3></div>
          <div style={{ width: "100%", height: 260 }}>
            <ResponsiveContainer>
              <BarChart data={countByProject} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A2F37" vertical={false} />
                <XAxis dataKey="name" stroke="#8B93A1" fontSize={12} tickLine={false} axisLine={{ stroke: "#333941" }} />
                <YAxis stroke="#8B93A1" fontSize={11} tickLine={false} axisLine={false} allowDecimals={false} />
                <Tooltip contentStyle={{ background: "#1D2127", border: "1px solid #333941", borderRadius: 8, fontSize: 12.5 }} />
                <Bar dataKey="equipment" fill="#FF7A45" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="sc-panel">
        <div className="sc-panel-head"><h3>Breakdown activity — last 14 days</h3><span className="sc-muted">based on status entries logged</span></div>
        <div style={{ width: "100%", height: 240 }}>
          <ResponsiveContainer>
            <LineChart data={last14} margin={{ top: 8, right: 12, left: -12, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2A2F37" vertical={false} />
              <XAxis dataKey="date" stroke="#8B93A1" fontSize={11.5} tickLine={false} axisLine={{ stroke: "#333941" }} />
              <YAxis stroke="#8B93A1" fontSize={11} tickLine={false} axisLine={false} allowDecimals={false} />
              <Tooltip contentStyle={{ background: "#1D2127", border: "1px solid #333941", borderRadius: 8, fontSize: 12.5 }} />
              <Legend wrapperStyle={{ fontSize: 12.5 }} />
              <Line type="monotone" dataKey="logged" name="Entries logged" stroke="#4A90D9" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="breakdown" name="Breakdown entries" stroke="#E5544B" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="sc-grid-2">
        <div className="sc-panel">
          <div className="sc-panel-head"><h3>Most frequent breakdowns</h3></div>
          <div className="sc-list">
            {downtimeLeaders.length === 0 && <div className="sc-empty">No breakdown entries logged yet.</div>}
            {downtimeLeaders.map(({ eq, n }) => (
              <div key={eq.id} className="sc-list-row">
                <div><div className="sc-strong">{eq.type}</div><div className="sc-muted sc-mono" style={{ fontSize: 12 }}>{eq.vehicleNo || eq.id} · {eq.project}</div></div>
                <div className="sc-tpi-chip sc-tpi-chip-overdue">{n} time{n === 1 ? "" : "s"}</div>
              </div>
            ))}
          </div>
        </div>
        <div className="sc-panel">
          <div className="sc-panel-head"><h3>Fleet composition</h3></div>
          <div className="sc-list">
            {typeBreakdown.map((t) => (
              <div key={t.type} className="sc-list-row">
                <div className="sc-strong">{t.type}</div>
                <div className="sc-bar-mini-wrap"><div className="sc-bar-mini" style={{ width: `${(t.count / typeBreakdown[0].count) * 100}%` }} /><span className="sc-mono" style={{ fontSize: 12.5 }}>{t.count}</span></div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ==================================================================== */
/* Styles                                                               */
/* ==================================================================== */

function Style() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

      .sc-root {
        --bg: #14171B; --panel: #1B1F25; --panel-2: #21262D; --border: #2C323A;
        --text: #E7E9EC; --muted: #8B93A1; --accent: #FF7A45; --blue: #4A90D9;
        font-family: 'Inter', system-ui, sans-serif; color: var(--text); background: var(--bg);
        min-height: 100%; width: 100%; box-sizing: border-box;
      }
      .sc-root * { box-sizing: border-box; }
      .sc-mono { font-family: 'IBM Plex Mono', monospace; }

      .sc-loading { display: flex; align-items: center; justify-content: center; min-height: 480px; }
      .sc-loading-inner { display: flex; align-items: center; gap: 10px; color: var(--muted); font-size: 14px; }
      .sc-spin { animation: sc-spin 1s linear infinite; }
      @keyframes sc-spin { to { transform: rotate(360deg); } }

      /* ---------------- Auth ---------------- */
      .sc-auth-root { min-height: 100vh; display: flex; align-items: center; justify-content: center; padding: 24px;
        background: radial-gradient(circle at 30% 20%, rgba(255,122,69,0.10), transparent 45%), radial-gradient(circle at 80% 80%, rgba(74,144,217,0.10), transparent 45%), var(--bg); }
      .sc-auth-card { width: 100%; max-width: 400px; background: var(--panel); border: 1px solid var(--border); border-radius: 16px; padding: 28px 26px; box-shadow: 0 24px 60px rgba(0,0,0,0.4); }
      .sc-auth-brand { display: flex; align-items: center; gap: 12px; margin-bottom: 22px; }
      .sc-auth-brand-name { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 19px; letter-spacing: 0.3px; }
      .sc-auth-brand-sub { font-size: 12.5px; color: var(--muted); margin-top: 1px; }
      .sc-auth-tabs { margin-bottom: 18px; }
      .sc-auth-foot { display: flex; gap: 8px; align-items: flex-start; margin-top: 18px; font-size: 11.5px; color: var(--muted); line-height: 1.5; }
      .sc-btn-block { width: 100%; justify-content: center; margin-top: 4px; }

      /* ---------------- Shell ---------------- */
      .sc-shell { display: flex; flex-direction: column; min-height: 100vh; }

      .sc-topbar {
        position: sticky; top: 0; z-index: 40; display: flex; align-items: center; gap: 18px;
        background: var(--panel); border-bottom: 1px solid var(--border); padding: 10px 22px;
      }
      .sc-topbar-left { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
      .sc-topbar-brand-text { line-height: 1.2; }
      .sc-brand-name { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 15px; letter-spacing: 0.2px; }
      .sc-brand-sub { font-size: 10.5px; color: var(--muted); margin-top: 1px; }

      .sc-topnav { display: flex; align-items: center; gap: 2px; flex: 1; overflow-x: auto; scrollbar-width: none; }
      .sc-topnav::-webkit-scrollbar { display: none; }
      .sc-topnav-item {
        display: flex; align-items: center; gap: 7px; padding: 8px 13px; border-radius: 8px;
        background: none; border: none; color: var(--muted); font-size: 13px; font-weight: 600;
        cursor: pointer; white-space: nowrap; font-family: inherit; flex-shrink: 0;
      }
      .sc-topnav-item:hover { background: var(--panel-2); color: var(--text); }
      .sc-topnav-item-active { background: var(--panel-2); color: var(--text); }
      .sc-topnav-item-active svg { color: var(--accent); }

      .sc-topbar-right { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
      .sc-segmented { display: flex; background: var(--panel-2); border: 1px solid var(--border); border-radius: 8px; padding: 2px; gap: 2px; }
      .sc-segmented-opt { flex: 1; background: none; border: none; color: var(--muted); font-size: 11.5px; font-weight: 600; padding: 6px 4px; border-radius: 6px; cursor: pointer; font-family: inherit; white-space: nowrap; }
      .sc-segmented-opt-active { background: var(--panel); color: var(--text); box-shadow: 0 1px 0 rgba(0,0,0,0.2); }
      .sc-segmented-compact { padding: 2px; }
      .sc-segmented-compact .sc-segmented-opt { padding: 5px 9px; font-size: 11px; }

      .sc-sync-dot { display: flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 7px; background: var(--panel-2); border: 1px solid var(--border); flex-shrink: 0; }

      .sc-user-chip { display: flex; align-items: center; gap: 8px; padding: 4px 6px 4px 4px; border-radius: 999px; background: var(--panel-2); border: 1px solid var(--border); }
      .sc-user-avatar { width: 26px; height: 26px; border-radius: 999px; background: var(--panel); display: flex; align-items: center; justify-content: center; color: var(--muted); flex-shrink: 0; }
      .sc-user-chip-text { min-width: 0; max-width: 140px; }
      .sc-user-chip-text > div { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

      .sc-hide-narrow { display: inline; }

      .sc-main { flex: 1; padding: 28px 34px 60px; min-width: 0; }
      .sc-page { display: flex; flex-direction: column; gap: 20px; max-width: 1180px; }

      .sc-page-head { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
      .sc-page-head h1 { font-family: 'Space Grotesk', sans-serif; font-size: 22px; font-weight: 700; margin: 0 0 4px; }
      .sc-page-head p { color: var(--muted); font-size: 13px; margin: 0; }

      /* ---------------- Hero ---------------- */
      .sc-hero { position: relative; border-radius: 16px; overflow: hidden; background: linear-gradient(135deg, #1E2229, #191C21); border: 1px solid var(--border); padding: 22px 26px; }
      .sc-hero-glow { position: absolute; top: -60px; right: -60px; width: 220px; height: 220px; border-radius: 50%; background: radial-gradient(circle, rgba(255,122,69,0.28), transparent 70%); pointer-events: none; }
      .sc-hero-inner { position: relative; display: flex; align-items: center; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
      .sc-hero-brand { display: flex; align-items: center; gap: 14px; }
      .sc-hero-title { font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 19px; }
      .sc-hero-sub { font-size: 12.5px; color: var(--muted); margin-top: 3px; }
      .sc-hero-date { font-size: 12.5px; color: var(--muted); background: rgba(255,255,255,0.04); border: 1px solid var(--border); padding: 6px 12px; border-radius: 999px; }

      .sc-stat-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; }
      .sc-stat { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 14px 15px; display: flex; flex-direction: column; gap: 8px; text-align: left; font-family: inherit; color: inherit; cursor: default; }
      .sc-stat-clickable { cursor: pointer; transition: border-color 0.15s, transform 0.1s; }
      .sc-stat-clickable:hover { border-color: var(--accent); transform: translateY(-1px); }
      .sc-stat-top { display: flex; align-items: center; justify-content: space-between; }
      .sc-stat-label { font-size: 11.5px; color: var(--muted); }
      .sc-stat-icon { width: 26px; height: 26px; border-radius: 7px; display: flex; align-items: center; justify-content: center; }
      .sc-stat-value { font-family: 'Space Grotesk', sans-serif; font-size: 25px; font-weight: 700; }
      .sc-stat-sub { font-size: 11px; color: var(--muted); }
      .sc-stat-cta { font-size: 11px; color: var(--accent); display: flex; align-items: center; gap: 2px; font-weight: 600; }

      .sc-grid-2 { display: grid; grid-template-columns: 1.3fr 1fr; gap: 16px; align-items: start; }
      .sc-panel { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 18px 20px; }
      .sc-panel-head { display: flex; align-items: baseline; justify-content: space-between; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; }
      .sc-panel-head h3 { font-family: 'Space Grotesk', sans-serif; font-size: 15px; font-weight: 600; margin: 0; }
      .sc-panel-head .sc-muted { font-size: 12px; }

      .sc-table-wrap { overflow-x: auto; margin: 0 -4px; }
      .sc-table { width: 100%; border-collapse: collapse; font-size: 13px; }
      .sc-table th { text-align: left; font-size: 11px; color: var(--muted); font-weight: 500; padding: 7px 10px; border-bottom: 1px solid var(--border); white-space: nowrap; }
      .sc-table td { padding: 9px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; white-space: nowrap; }
      .sc-table tr:last-child td { border-bottom: none; }
      .sc-table .sc-num { text-align: right; }
      .sc-strong { font-weight: 600; }
      .sc-muted { color: var(--muted); }
      .sc-dot { color: var(--border); }

      .sc-empty { color: var(--muted); font-size: 13px; padding: 18px 4px; text-align: center; }

      .sc-list { display: flex; flex-direction: column; gap: 2px; }
      .sc-list-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 9px 4px; border-bottom: 1px solid var(--border); }
      .sc-list-row:last-child { border-bottom: none; }

      .sc-tpi-chip { font-size: 11.5px; font-weight: 600; padding: 3px 9px; border-radius: 999px; background: rgba(227,164,56,0.14); color: #E3A438; white-space: nowrap; }
      .sc-tpi-chip-overdue { background: rgba(229,84,75,0.14); color: #E5544B; }

      .sc-chip-row { display: flex; flex-wrap: wrap; gap: 7px; }
      .sc-chip { font-size: 12px; padding: 5px 10px; border-radius: 7px; background: var(--panel-2); border: 1px solid var(--border); }
      .sc-chip-muted { color: var(--muted); }

      .sc-badge { display: inline-flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 600; padding: 4px 9px; border-radius: 999px; white-space: nowrap; }
      .sc-badge-sm { font-size: 11px; padding: 3px 8px; }
      .sc-rentedby-siscol { font-size: 11.5px; font-weight: 600; color: var(--accent); }
      .sc-rentedby-other { font-size: 11.5px; font-weight: 500; color: var(--muted); padding: 2px 8px; border: 1px solid var(--border); border-radius: 999px; }

      .sc-toolbar { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
      .sc-search { display: flex; align-items: center; gap: 8px; background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 8px 12px; flex: 1; min-width: 220px; color: var(--muted); }
      .sc-search input { background: none; border: none; outline: none; color: var(--text); font-size: 13px; width: 100%; font-family: inherit; }
      .sc-select { position: relative; display: flex; align-items: center; }
      .sc-select select { appearance: none; background: var(--panel); border: 1px solid var(--border); border-radius: 8px; padding: 8px 30px 8px 12px; color: var(--text); font-size: 13px; font-family: inherit; cursor: pointer; }
      .sc-select svg { position: absolute; right: 10px; pointer-events: none; color: var(--muted); }
      .sc-select-disabled select { opacity: 0.55; cursor: not-allowed; }

      .sc-cardlist { display: flex; flex-direction: column; gap: 8px; }
      .sc-eqcard { display: flex; align-items: center; justify-content: space-between; gap: 14px; background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 13px 16px; }
      .sc-eqcard-pending { border-left: 3px solid var(--accent); }
      .sc-eqcard-title { display: flex; align-items: center; gap: 8px; margin-bottom: 3px; }
      .sc-eqcard-meta { display: flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--muted); flex-wrap: wrap; }
      .sc-eqcard-side { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }

      .sc-btn { display: inline-flex; align-items: center; gap: 6px; font-family: inherit; font-size: 13px; font-weight: 600; padding: 8px 14px; border-radius: 8px; border: 1px solid transparent; cursor: pointer; }
      .sc-btn-primary { background: var(--accent); color: #17110C; }
      .sc-btn-primary:hover { background: #FF8A5C; }
      .sc-btn-ghost { background: var(--panel-2); color: var(--text); border-color: var(--border); }
      .sc-btn-ghost:hover { background: #282E36; }
      .sc-btn-danger { background: #E5544B; color: #fff; }

      .sc-icon-btn { background: none; border: none; color: var(--muted); cursor: pointer; padding: 6px; border-radius: 6px; display: inline-flex; }
      .sc-icon-btn:hover { background: var(--panel-2); color: var(--text); }
      .sc-icon-btn-danger:hover { color: #E5544B; }
      .sc-row-actions { display: flex; gap: 2px; }

      .sc-modal-overlay { position: fixed; inset: 0; background: rgba(10,12,14,0.6); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 20px; }
      .sc-modal { background: var(--panel); border: 1px solid var(--border); border-radius: 14px; width: 100%; max-width: 440px; max-height: 88vh; overflow-y: auto; font-family: 'Inter', sans-serif; color: var(--text); }
      .sc-modal-wide { max-width: 680px; }
      .sc-modal-head { display: flex; align-items: flex-start; justify-content: space-between; padding: 16px 20px; border-bottom: 1px solid var(--border); position: sticky; top: 0; background: var(--panel); }
      .sc-modal-head h3 { font-family: 'Space Grotesk', sans-serif; font-size: 15.5px; font-weight: 600; margin: 0; }
      .sc-modal-head p { font-size: 12px; color: var(--muted); margin: 3px 0 0; }
      .sc-modal-body { padding: 20px; }

      .sc-form { display: flex; flex-direction: column; gap: 14px; }
      .sc-form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
      .sc-field { display: flex; flex-direction: column; gap: 6px; font-size: 12.5px; }
      .sc-field-label { color: var(--muted); font-weight: 500; }
      .sc-field-hint { color: var(--muted); font-size: 11px; }
      .sc-input, .sc-textarea { background: var(--panel-2); border: 1px solid var(--border); border-radius: 7px; padding: 9px 11px; color: var(--text); font-size: 13.5px; font-family: inherit; outline: none; width: 100%; }
      .sc-input:focus, .sc-textarea:focus { border-color: var(--blue); }
      .sc-textarea { resize: vertical; }
      .sc-form-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 4px; padding-top: 14px; }

      .sc-status-picker { display: flex; flex-wrap: wrap; gap: 7px; }
      .sc-status-opt { display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 500; padding: 7px 11px; border-radius: 7px; border: 1px solid var(--border); background: var(--panel-2); color: var(--muted); cursor: pointer; font-family: inherit; }
      .sc-status-opt-active { font-weight: 600; }

      .sc-error { color: #E5544B; font-size: 12.5px; }

      .sc-project-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(230px, 1fr)); gap: 14px; }
      .sc-project-card { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; text-align: left; font-family: inherit; color: inherit; }
      .sc-project-card-clickable { cursor: pointer; transition: border-color 0.15s, transform 0.1s; }
      .sc-project-card-clickable:hover { border-color: var(--accent); transform: translateY(-1px); }
      .sc-project-card-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
      .sc-project-card-head h3 { font-family: 'Space Grotesk', sans-serif; font-size: 16px; font-weight: 700; margin: 0; }
      .sc-project-stats { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 14px; }
      .sc-project-stats > div { display: flex; flex-direction: column; gap: 3px; font-size: 13px; }
      .sc-project-stats .sc-muted { font-size: 11px; }
      .sc-project-card-cta { margin-top: 12px; font-size: 11.5px; color: var(--accent); font-weight: 600; display: flex; align-items: center; gap: 2px; }

      .sc-detail-list { display: flex; flex-direction: column; gap: 10px; }
      .sc-detail-card { background: var(--panel-2); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; }
      .sc-detail-card-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 10px; margin-bottom: 10px; }
      .sc-detail-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 16px; font-size: 12.5px; }
      .sc-detail-grid > div { display: flex; justify-content: space-between; gap: 8px; }
      .sc-detail-grid .sc-muted { font-size: 11.5px; }
      .sc-detail-note { margin-top: 10px; padding-top: 10px; border-top: 1px dashed var(--border); font-size: 12.5px; line-height: 1.5; }

      .sc-bar-mini-wrap { flex: 1; display: flex; align-items: center; gap: 8px; max-width: 200px; margin-left: 14px; }
      .sc-bar-mini { height: 7px; border-radius: 4px; background: var(--blue); min-width: 4px; }

      .sc-toast { position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%); background: var(--panel-2); border: 1px solid var(--border); color: var(--text); padding: 10px 18px; border-radius: 999px; font-size: 13px; box-shadow: 0 8px 24px rgba(0,0,0,0.35); z-index: 60; }

      .sc-export-btn { display: inline-flex; align-items: center; justify-content: center; gap: 7px; padding: 6px 11px; border-radius: 8px; border: 1px solid var(--border); background: var(--panel-2); color: var(--text); font-size: 12px; font-weight: 600; cursor: pointer; font-family: inherit; flex-shrink: 0; }
      .sc-export-btn:hover { border-color: var(--accent); color: var(--accent); }
      @media (max-width: 1600px) {
        .sc-hide-mid { display: none !important; }
        .sc-topnav-item { padding: 8px 10px; }
      }
      @media (max-width: 1400px) {
        .sc-topnav-item span:not(.sc-nav-badge) { display: none; }
        .sc-topnav-item-active span:not(.sc-nav-badge) { display: inline; }
        .sc-topnav-item { padding: 8px 10px; }
      }
      @media (max-width: 880px) {
        .sc-stat-grid { grid-template-columns: repeat(2, 1fr); }
        .sc-grid-2 { grid-template-columns: 1fr; }
        .sc-form-grid { grid-template-columns: 1fr; }
        .sc-topbar { padding: 8px 12px; gap: 8px 10px; flex-wrap: wrap; }
        .sc-topnav { order: 3; flex: 1 1 100%; }
        .sc-topbar-right { margin-left: auto; }
        .sc-topbar-brand-text { display: none; }
        .sc-hide-narrow { display: none; }
        .sc-main { padding: 20px 16px 60px; }
        .sc-eqcard { flex-direction: column; align-items: flex-start; }
        .sc-eqcard-side { width: 100%; justify-content: space-between; }
        .sc-hero-inner { flex-direction: column; align-items: flex-start; }
      }
    `}</style>
  );
}
