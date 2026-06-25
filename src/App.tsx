import { type FormEvent, useEffect, useMemo, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import ExcelJS from "exceljs";

type SaarathiRow = Record<string, string | number | boolean | null>;
type FilterKey = "Month" | "Subzone" | "RSM" | "ASM" | "SO" | "DB Name";
type Filters = Record<FilterKey, string>;
type QuickView = "all" | "unbilled" | "billedNoClub" | "inClubNoGrowth" | "growthNotInClub";
type SortDirection = "asc" | "desc";
type SortConfig = { column: string; direction: SortDirection };
type AuthView = "login" | "signup" | "forgot" | "reset";
type UserAccess = {
  id: number;
  email: string | null;
  role: string | null;
  rsm: string | null;
  asm: string | null;
  subzone: string | null;
  can_redeem: boolean | null;
  can_submit_redemption: boolean | null;
  is_active: boolean | null;
};
type OutletKpiBucket = {
  clubs: Set<string>;
  hasAchPositive: boolean;
  hasAchZero: boolean;
  achTotal: number;
  payoutTotal: number;
  growthTotal: number;
  paymentIdTotal: number;
  noOfSkuTotal: number;
};

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL as string;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string;
const tableName = (import.meta.env.VITE_SUPABASE_TABLE as string) || "Saarathi";

const supabase = createClient(supabaseUrl, supabaseAnonKey);
const FILTER_KEYS: FilterKey[] = ["Month", "Subzone", "RSM", "ASM", "SO", "DB Name"];
const HIERARCHY_KEYS: FilterKey[] = ["Subzone", "RSM", "ASM", "SO", "DB Name"];

const EMPTY_FILTERS: Filters = {
  Month: "",
  Subzone: "",
  RSM: "",
  ASM: "",
  SO: "",
  "DB Name": "",
};
const PAGE_SIZE = 1000;
const TABLE_PAGE_SIZE = 20;
const MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const DISPLAY_COLUMNS = [
  "SO",
  "ASM",
  "RSM",
  "DB Code",
  "DB Name",
  "Outlet Id",
  "Outlet Name",
  "LYSM Tgt",
  "Ach",
  "Club",
  "Growth %",
  "LPPC",
  "Payout %",
] as const;
const SEARCH_COLUMNS = ["SO", "ASM", "RSM", "DB Code", "DB Name", "Outlet Id", "Outlet Name", "Club"] as const;
const EXPORT_COLUMNS = [
  { key: "SO", label: "SO Name" },
  { key: "ASM", label: "ASM" },
  { key: "RSM", label: "RSM" },
  { key: "DB Code", label: "DB Code" },
  { key: "DB Name", label: "DB Name" },
  { key: "Outlet Id", label: "Outlet ID" },
  { key: "Outlet Name", label: "Outlet Name" },
  { key: "LYSM Tgt", label: "LYSM" },
  { key: "Ach", label: "Ach" },
  { key: "Club", label: "Club" },
  { key: "Growth %", label: "Growth %" },
  { key: "LPPC", label: "LPPC" },
  { key: "Payout %", label: "Payout %" },
  { key: "Beat Id", label: "Beat Id" },
  { key: "Beat Name", label: "Beat Name" },
  { key: "Order Value", label: "Order Value" },
] as const;

const getCellValue = (row: SaarathiRow, key: FilterKey): string => {
  const value = row[key];
  return value === null || value === undefined ? "" : String(value).trim();
};

const parseNumber = (value: unknown): number => {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/,/g, "").replace(/%/g, "").trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeClub = (value: unknown): string => {
  return String(value ?? "").trim().toLowerCase();
};

const getCurrentMonthLabel = (): string => {
  const now = new Date();
  const month = MONTH_ORDER[now.getMonth()];
  const year = String(now.getFullYear()).slice(-2);
  return `${month}-${year}`;
};

const getMonthSortValue = (monthLabel: string): number => {
  const cleaned = monthLabel.trim();
  const [mon, yy] = cleaned.split("-");
  if (!mon || !yy) return Number.MAX_SAFE_INTEGER;

  const monthIndex = MONTH_ORDER.indexOf(mon);
  const year = Number(yy);

  if (monthIndex === -1 || Number.isNaN(year)) return Number.MAX_SAFE_INTEGER;
  return year * 12 + monthIndex;
};

const formatRounded = (value: unknown): string => {
  return Math.round(parseNumber(value)).toLocaleString("en-IN");
};

const formatTotalValue = (value: number): string => {
  if (value >= 100000) {
    const lakhValue = value / 100000;
    return `${lakhValue.toFixed(2)} L`;
  }
  return Math.round(value).toLocaleString("en-IN");
};

const parsePercent = (value: unknown): number => {
  const raw = parseNumber(value);
  if (!Number.isFinite(raw)) return 0;
  return Math.max(0, Math.min(100, raw));
};

const getRowLppc = (row: SaarathiRow): number => {
  const paymentIdTotal = parseNumber(row["Payment id"]);
  const noOfSkuTotal = parseNumber(row["No of Sku"]);
  return paymentIdTotal > 0 ? noOfSkuTotal / paymentIdTotal : 0;
};

const normalizeText = (value: unknown): string => String(value ?? "").trim().toLowerCase();

const matchesAccessList = (rowValue: unknown, accessValue: unknown): boolean => {
  const rowNorm = normalizeText(rowValue);
  const accessNorm = normalizeText(accessValue);
  if (!accessNorm) return false;

  const allowed = accessNorm
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  return allowed.includes(rowNorm);
};

export default function App() {
  const [rows, setRows] = useState<SaarathiRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [signupEmail, setSignupEmail] = useState("");
  const [signupPassword, setSignupPassword] = useState("");
  const [resetPassword, setResetPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [authView, setAuthView] = useState<AuthView>("login");
  const [loginError, setLoginError] = useState<string | null>(null);
  const [authMessage, setAuthMessage] = useState<string | null>(null);
  const [accessRow, setAccessRow] = useState<UserAccess | null>(null);
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [quickView, setQuickView] = useState<QuickView>("all");
  const [sortConfig, setSortConfig] = useState<SortConfig>({ column: "Ach", direction: "desc" });
  const [currentPage, setCurrentPage] = useState(1);
  const [searchText, setSearchText] = useState("");
  const [filtersOpenMobile, setFiltersOpenMobile] = useState(false);

  useEffect(() => {
    const initAuth = async () => {
      const { data } = await supabase.auth.getSession();
      const email = data.session?.user?.email ?? null;
      setUserEmail(email);
      setAuthLoading(false);
    };

    void initAuth();

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      if (_event === "PASSWORD_RECOVERY") {
        setAuthView("reset");
      }
      setUserEmail(session?.user?.email ?? null);
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    const fetchAccess = async () => {
      if (!userEmail) {
        setAccessRow(null);
        return;
      }

      const normalizedEmail = userEmail.trim().toLowerCase();
      const { data, error: accessError } = await supabase
        .schema("public")
        .from("user_access_control")
        .select("*")
        .ilike("email", normalizedEmail)
        .maybeSingle();

      if (accessError) {
        setAccessRow(null);
        setError(`Access check failed: ${accessError.message}`);
        return;
      }

      if (!data) {
        const { data: inserted, error: insertError } = await supabase
          .schema("public")
          .from("user_access_control")
          .insert({
            email: normalizedEmail,
            is_active: true,
            can_redeem: true,
            can_submit_redemption: false,
          })
          .select("*")
          .single();

        if (insertError) {
          setAccessRow(null);
          setError(`Access row creation failed: ${insertError.message}`);
          return;
        }

        setAccessRow(inserted as UserAccess);
        return;
      }

      setAccessRow((data as UserAccess).is_active ? (data as UserAccess) : null);
    };

    void fetchAccess();
  }, [userEmail]);

  useEffect(() => {
    let active = true;

    const loadRows = async () => {
      setLoading(true);
      setError(null);

      const allRows: SaarathiRow[] = [];
      let from = 0;

      while (true) {
        const to = from + PAGE_SIZE - 1;
        const { data, error: fetchError } = await supabase
          .schema("public")
          .from(tableName)
          .select("*")
          .range(from, to);

        if (!active) return;

        if (fetchError) {
          setError(fetchError.message);
          setRows([]);
          setLoading(false);
          return;
        }

        const batch = (data ?? []) as SaarathiRow[];
        allRows.push(...batch);

        if (batch.length < PAGE_SIZE) {
          break;
        }

        from += PAGE_SIZE;
      }

      if (!active) return;
      setRows(allRows);

      const monthOptions = Array.from(
        new Set(
          allRows
            .map((row) => getCellValue(row, "Month"))
            .filter((value) => value !== ""),
        ),
      ).sort((a, b) => getMonthSortValue(a) - getMonthSortValue(b));

      if (monthOptions.length > 0) {
        const currentMonth = getCurrentMonthLabel();
        const defaultMonth = monthOptions.includes(currentMonth)
          ? currentMonth
          : monthOptions[monthOptions.length - 1];

        setFilters((prev) => (prev.Month ? prev : { ...prev, Month: defaultMonth }));
      }

      setLoading(false);
    };

    void loadRows();

    return () => {
      active = false;
    };
  }, []);

  const handleLogin = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setLoginError(null);
    setAuthMessage(null);
    const { error: signInError } = await supabase.auth.signInWithPassword({
      email: loginEmail.trim(),
      password: loginPassword,
    });
    if (signInError) {
      setLoginError(signInError.message);
    }
  };

  const handleSignup = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setLoginError(null);
    setAuthMessage(null);
    const cleanEmail = signupEmail.trim().toLowerCase();
    const { error: signUpError } = await supabase.auth.signUp({
      email: cleanEmail,
      password: signupPassword,
      options: {
        emailRedirectTo: window.location.origin,
      },
    });

    if (signUpError) {
      const message = signUpError.message ?? "";
      if (message.toLowerCase().includes("already registered")) {
        setAuthMessage("User already registered. Please login with your password.");
        setLoginEmail(cleanEmail);
        setAuthView("login");
        return;
      }
      if (message.toLowerCase().includes("confirmation email")) {
        const { error: tryLoginError } = await supabase.auth.signInWithPassword({
          email: cleanEmail,
          password: signupPassword,
        });

        if (tryLoginError) {
          setLoginError(
            "Email confirmation is enabled in Supabase. Disable it in Authentication > Providers > Email, then try signup again.",
          );
          return;
        }
      } else {
        setLoginError(signUpError.message);
        return;
      }
    }

    const { data: existingAccess } = await supabase
      .schema("public")
      .from("user_access_control")
      .select("id")
      .eq("email", cleanEmail)
      .maybeSingle();

    const accessInsertNeeded = !existingAccess?.id;
    const { error: accessInsertError } = accessInsertNeeded
      ? await supabase
          .schema("public")
          .from("user_access_control")
          .insert({
            email: cleanEmail,
            is_active: true,
            can_redeem: true,
            can_submit_redemption: false,
          })
      : { error: null };

    if (accessInsertError) {
      // Continue, because auth signup already succeeded; admin can still add row manually.
      setAuthMessage(
        "Signup created. Access row insert failed, please ask admin to map your role/subzone/asm/rsm.",
      );
      setAuthView("login");
      return;
    }

    setAuthMessage("Signup successful. You can login now.");
    setAuthView("login");
  };

  const handleForgotPassword = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setLoginError(null);
    setAuthMessage(null);
    const { error: resetError } = await supabase.auth.resetPasswordForEmail(loginEmail.trim(), {
      redirectTo: window.location.origin,
    });
    if (resetError) {
      setLoginError(resetError.message);
      return;
    }
    setAuthMessage("Password reset link sent to your email.");
  };

  const handleResetPassword = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setLoginError(null);
    setAuthMessage(null);
    const { error: updateError } = await supabase.auth.updateUser({ password: resetPassword });
    if (updateError) {
      setLoginError(updateError.message);
      return;
    }
    setAuthMessage("Password updated successfully. Please login.");
    setAuthView("login");
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setAccessRow(null);
    setFilters(EMPTY_FILTERS);
    setQuickView("all");
    setSearchText("");
    setCurrentPage(1);
  };

  const accessibleRows = useMemo(() => {
    if (!accessRow) return [];
    const role = String(accessRow.role ?? "").trim().toLowerCase();

    if (role === "admin") return rows;
    if (role === "rsm") {
      return rows.filter((row) => {
        if (accessRow.subzone && matchesAccessList(row.Subzone, accessRow.subzone)) return true;
        if (accessRow.rsm && matchesAccessList(row.RSM, accessRow.rsm)) return true;
        return false;
      });
    }
    if (role === "asm") {
      return rows.filter((row) => (accessRow.asm ? matchesAccessList(row.ASM, accessRow.asm) : false));
    }
    return [];
  }, [rows, accessRow]);

  const columns = DISPLAY_COLUMNS.filter((column) =>
    column === "LPPC"
      ? accessibleRows.some(
          (row) => Object.prototype.hasOwnProperty.call(row, "Payment id") || Object.prototype.hasOwnProperty.call(row, "No of Sku"),
        )
      : accessibleRows.some((row) => Object.prototype.hasOwnProperty.call(row, column)),
  );

  const rowsByHierarchy = useMemo(() => {
    return accessibleRows.filter((row) => {
      if (filters.Month && getCellValue(row, "Month") !== filters.Month) {
        return false;
      }

      return HIERARCHY_KEYS.every((key) => {
        const selected = filters[key];
        if (!selected) return true;
        return getCellValue(row, key) === selected;
      });
    });
  }, [accessibleRows, filters]);

  const optionsByFilter = useMemo(() => {
    const result = {} as Record<FilterKey, string[]>;
    for (const key of FILTER_KEYS) {
      const values = accessibleRows
        .filter((row) => {
          if (filters.Month && key !== "Month" && getCellValue(row, "Month") !== filters.Month) {
            return false;
          }

          if (key === "Month") return true;

          const currentIndex = HIERARCHY_KEYS.indexOf(key);
          if (currentIndex === -1) return true;

          for (let i = 0; i < currentIndex; i += 1) {
            const parentKey = HIERARCHY_KEYS[i];
            const selectedParent = filters[parentKey];
            if (selectedParent && getCellValue(row, parentKey) !== selectedParent) {
              return false;
            }
          }

          return true;
        })
        .map((row) => getCellValue(row, key))
        .filter((value) => value !== "");

      const uniqueValues = Array.from(new Set(values));
      result[key] =
        key === "Month"
          ? uniqueValues.sort((a, b) => getMonthSortValue(a) - getMonthSortValue(b))
          : uniqueValues.sort((a, b) => a.localeCompare(b));
    }
    return result;
  }, [accessibleRows, filters]);

  const handleFilterChange = (key: FilterKey, value: string) => {
    setCurrentPage(1);
    setFilters((prev) => {
      const next = { ...prev, [key]: value };

      if (key === "Month") {
        return next;
      }

      const changedIndex = HIERARCHY_KEYS.indexOf(key);
      if (changedIndex !== -1) {
        for (let i = changedIndex + 1; i < HIERARCHY_KEYS.length; i += 1) {
          next[HIERARCHY_KEYS[i]] = "";
        }
      }

      return next;
    });
  };

  const clearFilters = () => {
    setFilters(EMPTY_FILTERS);
    setQuickView("all");
    setSearchText("");
    setCurrentPage(1);
  };

  const dataAfterQuickView = useMemo(() => {
    if (quickView === "unbilled") {
      return rowsByHierarchy.filter((row) => parseNumber(row.Ach) === 0);
    }
    if (quickView === "billedNoClub") {
      return rowsByHierarchy.filter(
        (row) => parseNumber(row.Ach) > 0 && normalizeClub(row.Club) === "no club",
      );
    }
    if (quickView === "inClubNoGrowth") {
      return rowsByHierarchy.filter((row) => {
        const club = normalizeClub(row.Club);
        const inClub = club === "diamond" || club === "daimond" || club === "gold" || club === "silver";
        return inClub && parseNumber(row["Payout %"]) === 0;
      });
    }
    if (quickView === "growthNotInClub") {
      return rowsByHierarchy.filter(
        (row) => parseNumber(row["Growth %"]) >= 30 && normalizeClub(row.Club) === "no club",
      );
    }
    return rowsByHierarchy;
  }, [rowsByHierarchy, quickView]);

  const rowsAfterSearch = useMemo(() => {
    const q = searchText.trim().toLowerCase();
    if (!q) return dataAfterQuickView;
    return dataAfterQuickView.filter((row) =>
      SEARCH_COLUMNS.some((col) => String(row[col] ?? "").toLowerCase().includes(q)),
    );
  }, [dataAfterQuickView, searchText]);

  const sortedRows = useMemo(() => {
    const sorted = [...rowsAfterSearch];
    sorted.sort((a, b) => {
      const aValue = sortConfig.column === "LPPC" ? getRowLppc(a) : a[sortConfig.column];
      const bValue = sortConfig.column === "LPPC" ? getRowLppc(b) : b[sortConfig.column];
      const numericColumns = ["Outlet Id", "LYSM Tgt", "Ach", "Growth %", "LPPC", "Payout %"];

      if (numericColumns.includes(sortConfig.column)) {
        const diff = parseNumber(aValue) - parseNumber(bValue);
        return sortConfig.direction === "asc" ? diff : -diff;
      }

      const aText = String(aValue ?? "").toLowerCase();
      const bText = String(bValue ?? "").toLowerCase();
      const cmp = aText.localeCompare(bText);
      return sortConfig.direction === "asc" ? cmp : -cmp;
    });
    return sorted;
  }, [rowsAfterSearch, sortConfig]);

  const totalPages = Math.max(1, Math.ceil(sortedRows.length / TABLE_PAGE_SIZE));
  const safeCurrentPage = Math.min(currentPage, totalPages);
  const pagedRows = useMemo(() => {
    const start = (safeCurrentPage - 1) * TABLE_PAGE_SIZE;
    return sortedRows.slice(start, start + TABLE_PAGE_SIZE);
  }, [sortedRows, safeCurrentPage]);

  const handleSort = (column: string) => {
    setCurrentPage(1);
    setSortConfig((prev) => {
      if (prev.column === column) {
        return { column, direction: prev.direction === "desc" ? "asc" : "desc" };
      }
      return { column, direction: "desc" };
    });
  };

  const downloadExcel = () => {
    void (async () => {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Codex";
      workbook.lastModifiedBy = "Codex";
      workbook.created = new Date();
      workbook.modified = new Date();
      workbook.properties.date1904 = true;

      const sheet = workbook.addWorksheet("Saarathi", {
        views: [{ state: "frozen", ySplit: 1 }],
      });

      sheet.columns = EXPORT_COLUMNS.map(({ label }) => ({
        header: label,
        key: label,
        width:
          label === "DB Name" || label === "Outlet Name"
            ? 32
            : label === "SO Name" || label === "ASM" || label === "RSM"
              ? 22
              : label === "Beat Name"
                ? 20
                : label === "Order Value" || label === "LPPC"
                  ? 14
                  : 14,
      }));

      const headerRow = sheet.getRow(1);
      headerRow.height = 22;
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.alignment = { vertical: "middle", horizontal: "center" };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1F5B99" },
      };
      headerRow.border = {
        top: { style: "thin", color: { argb: "FFB7C9E2" } },
        left: { style: "thin", color: { argb: "FFB7C9E2" } },
        bottom: { style: "thin", color: { argb: "FFB7C9E2" } },
        right: { style: "thin", color: { argb: "FFB7C9E2" } },
      };

      for (const row of sortedRows) {
        const dataRow = EXPORT_COLUMNS.map(({ key }) => {
          if (key === "LYSM Tgt" || key === "Ach" || key === "Order Value" || key === "LPPC") {
            return key === "LPPC" ? Number(getRowLppc(row).toFixed(2)) : Math.round(parseNumber(row[key]));
          }
          if (key === "Outlet Id") {
            return Math.round(parseNumber(row[key]));
          }
          return row[key] === null || row[key] === undefined ? "" : String(row[key]);
        });
        const added = sheet.addRow(dataRow);
        added.height = 18;
        added.eachCell((cell, columnNumber) => {
          const label = EXPORT_COLUMNS[columnNumber - 1]?.label ?? "";
          const centerColumns = new Set(["Outlet ID", "LYSM", "Ach", "Growth %", "LPPC", "Payout %", "Order Value"]);
          cell.alignment = {
            vertical: "middle",
            horizontal: centerColumns.has(label) ? "center" : "left",
          };
          cell.border = {
            top: { style: "thin", color: { argb: "FFD9E2F3" } },
            left: { style: "thin", color: { argb: "FFD9E2F3" } },
            bottom: { style: "thin", color: { argb: "FFD9E2F3" } },
            right: { style: "thin", color: { argb: "FFD9E2F3" } },
          };
        });
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "levista_saarathi_club.xlsx";
      anchor.click();
      URL.revokeObjectURL(url);
    })();
  };

  const kpis = useMemo(() => {
    const outletBuckets = new Map<string, OutletKpiBucket>();
    for (const row of rowsByHierarchy) {
      const outletId = row["Outlet Id"];
      if (outletId !== null && outletId !== undefined && String(outletId).trim() !== "") {
        const outletKey = String(outletId);
        const club = normalizeClub(row.Club);
        const ach = parseNumber(row.Ach);
        const payout = parseNumber(row["Payout %"]);
        const growth = parseNumber(row["Growth %"]);

        const existing = outletBuckets.get(outletKey);
        if (!existing) {
          outletBuckets.set(outletKey, {
            clubs: new Set(club ? [club] : []),
            hasAchPositive: ach > 0,
            hasAchZero: ach <= 0,
            achTotal: ach,
            payoutTotal: payout,
            growthTotal: growth,
            paymentIdTotal: parseNumber(row["Payment id"]),
            noOfSkuTotal: parseNumber(row["No of Sku"]),
          });
        } else {
          if (club) existing.clubs.add(club);
          existing.hasAchPositive = existing.hasAchPositive || ach > 0;
          existing.hasAchZero = existing.hasAchZero || ach <= 0;
          existing.achTotal += ach;
          existing.payoutTotal += payout;
          existing.growthTotal += growth;
          existing.paymentIdTotal += parseNumber(row["Payment id"]);
          existing.noOfSkuTotal += parseNumber(row["No of Sku"]);
        }
      }
    }

    const uniqueOutlets = new Set(outletBuckets.keys());
    const diamondOutlets = new Set<string>();
    const goldOutlets = new Set<string>();
    const silverOutlets = new Set<string>();
    const noClubOutlets = new Set<string>();
    const coverageOutlets = new Set<string>();
    const growthOutlets = new Set<string>();
    const inClubNoGrowthOutlets = new Set<string>();
    const growthNotInClubOutlets = new Set<string>();
    const billedNotInClubOutlets = new Set<string>();
    let billedNotInClub = 0;
    let totalAch = 0;
    let paymentIdTotal = 0;
    let noOfSkuTotal = 0;

    for (const [outletKey, bucket] of outletBuckets) {
      if (bucket.clubs.has("diamond") || bucket.clubs.has("daimond")) diamondOutlets.add(outletKey);
      if (bucket.clubs.has("gold")) goldOutlets.add(outletKey);
      if (bucket.clubs.has("silver")) silverOutlets.add(outletKey);
      if (bucket.clubs.has("no club")) noClubOutlets.add(outletKey);

      if (bucket.hasAchPositive) coverageOutlets.add(outletKey);
      if (bucket.payoutTotal > 0) growthOutlets.add(outletKey);
      if (
        (bucket.clubs.has("diamond") ||
          bucket.clubs.has("daimond") ||
          bucket.clubs.has("gold") ||
          bucket.clubs.has("silver")) &&
        bucket.payoutTotal === 0
      ) {
        inClubNoGrowthOutlets.add(outletKey);
      }
      if (bucket.clubs.has("no club") && bucket.growthTotal >= 30) {
        growthNotInClubOutlets.add(outletKey);
      }

      totalAch += bucket.achTotal;
      paymentIdTotal += bucket.paymentIdTotal;
      noOfSkuTotal += bucket.noOfSkuTotal;
    }

    const unbilledOutlets = new Set<string>();
    for (const [outletKey, bucket] of outletBuckets) {
      if (bucket.achTotal <= 0) {
        unbilledOutlets.add(outletKey);
      }
      if (bucket.achTotal > 0 && bucket.clubs.has("no club")) {
        billedNotInClubOutlets.add(outletKey);
      }
    }

    const unbilled = unbilledOutlets.size;
    billedNotInClub = billedNotInClubOutlets.size;

    return {
      totalOutletsEnrolled: uniqueOutlets.size,
      coverageAchievement:
        uniqueOutlets.size > 0 ? (coverageOutlets.size / uniqueOutlets.size) * 100 : 0,
      coverageOutletCount: coverageOutlets.size,
      outletsGrowthAchievement:
        uniqueOutlets.size > 0 ? (growthOutlets.size / uniqueOutlets.size) * 100 : 0,
      growthOutletCount: growthOutlets.size,
      diamond: diamondOutlets.size,
      gold: goldOutlets.size,
      silver: silverOutlets.size,
      noClub: noClubOutlets.size,
      unbilled,
      billedNotInClub,
      inClubNoGrowth: inClubNoGrowthOutlets.size,
      growthNotInClub: growthNotInClubOutlets.size,
      totalAch,
      paymentIdTotal,
      noOfSkuTotal,
      lppc: paymentIdTotal > 0 ? noOfSkuTotal / paymentIdTotal : 0,
    };
  }, [rowsByHierarchy]);

  const pageList = useMemo(() => {
    const pages: (number | "...")[] = [];
    if (totalPages <= 7) {
      for (let p = 1; p <= totalPages; p += 1) pages.push(p);
      return pages;
    }
    pages.push(1);
    if (safeCurrentPage > 3) pages.push("...");
    const start = Math.max(2, safeCurrentPage - 1);
    const end = Math.min(totalPages - 1, safeCurrentPage + 1);
    for (let p = start; p <= end; p += 1) pages.push(p);
    if (safeCurrentPage < totalPages - 2) pages.push("...");
    pages.push(totalPages);
    return pages;
  }, [safeCurrentPage, totalPages]);

  if (authLoading) {
    return (
      <main className="container">
        <h1>Levista Saarathi Club</h1>
        <p>Checking login session...</p>
      </main>
    );
  }

  if (!userEmail) {
    return (
      <main className="container authScreen">
        <h1>Levista Saarathi Club</h1>
        {authView === "login" && (
          <form className="loginCard" onSubmit={handleLogin}>
            <h2>Login</h2>
            <label>
              Email
              <input
                type="email"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
                required
              />
            </label>
            <label>
              Password
              <input
                type={showPassword ? "text" : "password"}
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
                required
              />
            </label>
            <label className="checkRow">
              <input
                type="checkbox"
                checked={showPassword}
                onChange={(e) => setShowPassword(e.target.checked)}
              />
              Show password
            </label>
            {loginError ? <p className="error">{loginError}</p> : null}
            {authMessage ? <p className="successMsg">{authMessage}</p> : null}
            <button type="submit" className="loginBtn">
              Login
            </button>
          </form>
        )}

        {authView === "signup" && (
          <form className="loginCard" onSubmit={handleSignup}>
            <h2>First Time Signup</h2>
            <label>
              Email
              <input
                type="email"
                value={signupEmail}
                onChange={(e) => setSignupEmail(e.target.value)}
                required
              />
            </label>
            <label>
              Set Password
              <input
                type={showPassword ? "text" : "password"}
                value={signupPassword}
                onChange={(e) => setSignupPassword(e.target.value)}
                required
              />
            </label>
            <label className="checkRow">
              <input
                type="checkbox"
                checked={showPassword}
                onChange={(e) => setShowPassword(e.target.checked)}
              />
              Show password
            </label>
            {loginError ? <p className="error">{loginError}</p> : null}
            {authMessage ? <p className="successMsg">{authMessage}</p> : null}
            <button type="submit" className="loginBtn">
              Create Account
            </button>
            <button type="button" className="linkBtn" onClick={() => setAuthView("login")}>
              Back to login
            </button>
          </form>
        )}

        {authView === "forgot" && (
          <form className="loginCard" onSubmit={handleForgotPassword}>
            <h2>Forgot Password</h2>
            <label>
              Email
              <input
                type="email"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
                required
              />
            </label>
            {loginError ? <p className="error">{loginError}</p> : null}
            {authMessage ? <p className="successMsg">{authMessage}</p> : null}
            <button type="submit" className="loginBtn">
              Send Reset Link
            </button>
            <button type="button" className="linkBtn" onClick={() => setAuthView("login")}>
              Back to login
            </button>
          </form>
        )}

        {authView === "reset" && (
          <form className="loginCard" onSubmit={handleResetPassword}>
            <h2>Reset Password</h2>
            <label>
              New Password
              <input
                type={showPassword ? "text" : "password"}
                value={resetPassword}
                onChange={(e) => setResetPassword(e.target.value)}
                required
              />
            </label>
            <label className="checkRow">
              <input
                type="checkbox"
                checked={showPassword}
                onChange={(e) => setShowPassword(e.target.checked)}
              />
              Show password
            </label>
            {loginError ? <p className="error">{loginError}</p> : null}
            {authMessage ? <p className="successMsg">{authMessage}</p> : null}
            <button type="submit" className="loginBtn">
              Update Password
            </button>
          </form>
        )}
      </main>
    );
  }

  if (!accessRow) {
    return (
      <main className="container authScreen">
        <h1>Levista Saarathi Club</h1>
        <p className="error">No active access found for {userEmail}. Contact admin.</p>
        <button type="button" className="loginBtn" onClick={handleLogout}>
          Logout
        </button>
      </main>
    );
  }

  if (loading) {
    return (
      <main className="container">
        <h1>Levista Saarathi Club</h1>
        <p>Loading data...</p>
      </main>
    );
  }

  if (error) {
    return (
      <main className="container">
        <h1>Levista Saarathi Club</h1>
        <p className="meta">Tried table: public.{tableName}</p>
        <p className="error">Error loading data: {error}</p>
      </main>
    );
  }

  return (
    <main className="container">
      <div className="topHeader">
        <button
          type="button"
          className="mobileFilterToggle"
          onClick={() => setFiltersOpenMobile((prev) => !prev)}
        >
          {filtersOpenMobile ? "Hide Filters" : "Show Filters"}
        </button>

        <h1 className="appTitle">Levista Saarathi Club</h1>

        <div className="userBlock">
          <p className="userNameLabel">{userEmail}</p>
          <button type="button" className="logoutBtn" onClick={handleLogout}>
            Logout
          </button>
        </div>
      </div>

      <section className="kpiGrid">
        <article className="kpiCard kpiCardPrimary">
          <p className="kpiLabel">Total Outlets Enrolled</p>
          <p className="kpiValue">{kpis.totalOutletsEnrolled}</p>
          <p className="kpiLabel kpiLabelSpacing">Slabs</p>
          <p className="kpiSubline">Diamond: {kpis.diamond}</p>
          <p className="kpiSubline">Gold: {kpis.gold}</p>
          <p className="kpiSubline">Silver: {kpis.silver}</p>
          <p className="kpiSubline">No Club: {kpis.noClub}</p>
        </article>

        <article className="kpiCard kpiCardSmooth">
          <p className="kpiLabel">Coverage Target 98%</p>
          <p className="kpiValue">
            {kpis.coverageAchievement.toFixed(1)}% <span className="kpiCount">({kpis.coverageOutletCount})</span>
          </p>
          <p className="kpiLabel kpiLabelSpacing">Outlets Growth Target 60%</p>
          <p className="kpiValue">
            {kpis.outletsGrowthAchievement.toFixed(1)}% <span className="kpiCount">({kpis.growthOutletCount})</span>
          </p>
          <p className="kpiLabel kpiLabelSpacing">LPPC</p>
          <p className="kpiValue">{kpis.lppc.toFixed(2)}</p>
        </article>

        <article className="kpiCard">
          <p className="kpiLabel">Unbilled Outlets</p>
          <p className="kpiValue">{kpis.unbilled}</p>
          <p className="kpiLabel kpiLabelSpacing">Billed but Not in Club</p>
          <p className="kpiValue kpiValueSmall">{kpis.billedNotInClub}</p>
          <p className="kpiLabel kpiLabelSpacing">In Club But No Growth</p>
          <p className="kpiValue kpiValueSmall">{kpis.inClubNoGrowth}</p>
          <p className="kpiLabel kpiLabelSpacing">Growth But Not in Club</p>
          <p className="kpiValue kpiValueSmall">{kpis.growthNotInClub}</p>
        </article>

        <article className="kpiCard">
          <p className="kpiLabel">Total Value (Ach)</p>
          <p className="kpiValue">{formatTotalValue(kpis.totalAch)}</p>
        </article>
      </section>

      <div
        className={`filterBackdrop ${filtersOpenMobile ? "show" : ""}`}
        onClick={() => setFiltersOpenMobile(false)}
      />

      <div className="contentGrid">
        <aside className={`filtersPanel ${filtersOpenMobile ? "openMobile" : "closedMobile"}`}>
          <div className="filtersHeader">
            <h2>Filters</h2>
            <div className="filtersHeaderActions">
              <button type="button" className="clearBtn" onClick={clearFilters}>
                Clear
              </button>
              <button type="button" className="filtersCloseBtn" onClick={() => setFiltersOpenMobile(false)}>
                Close
              </button>
            </div>
          </div>
          {FILTER_KEYS.map((key) => (
            <label key={key} className="filterField">
              <span>{key}</span>
              <select
                value={filters[key]}
                onChange={(event) => handleFilterChange(key, event.target.value)}
              >
                <option value="">All</option>
                {optionsByFilter[key].map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
          ))}
        </aside>

        <div className="tableWrap">
          <div className="tableTopBar">
            <div />
            <div className="tableActions">
            <button
              type="button"
              className={`pillBtn ${quickView === "unbilled" ? "active" : ""}`}
              onClick={() => {
                setQuickView("unbilled");
                setCurrentPage(1);
              }}
            >
              Unbilled Outlets
            </button>
            <button
              type="button"
              className={`pillBtn ${quickView === "billedNoClub" ? "active" : ""}`}
              onClick={() => {
                setQuickView("billedNoClub");
                setCurrentPage(1);
              }}
            >
              Billed But Not in Club
            </button>
            <button
              type="button"
              className={`pillBtn ${quickView === "inClubNoGrowth" ? "active" : ""}`}
              onClick={() => {
                setQuickView("inClubNoGrowth");
                setCurrentPage(1);
              }}
            >
              In Club But No Growth
            </button>
            <button
              type="button"
              className={`pillBtn ${quickView === "growthNotInClub" ? "active" : ""}`}
              onClick={() => {
                setQuickView("growthNotInClub");
                setCurrentPage(1);
              }}
            >
              Growth But Not in Club
            </button>
            <button type="button" className="pillBtn download" onClick={downloadExcel}>
              Download
            </button>
          </div>
          </div>
          <div className="tableSearchBar">
            <input
              type="text"
              value={searchText}
              onChange={(e) => {
                setSearchText(e.target.value);
                setCurrentPage(1);
              }}
              placeholder="Search SO, ASM, RSM, DB Code, DB Name, Outlet ID, Outlet Name, Club"
            />
          </div>
          <table>
            <thead>
              <tr>
                {columns.map((column) => (
                  <th
                    key={column}
                    onClick={() => handleSort(column)}
                    className={`sortableHeader ${column === "SO" ? "stickyCol1" : ""} ${column === "ASM" ? "stickyCol2" : ""}`}
                  >
                    {column === "LYSM Tgt" ? "LYSM" : column}{" "}
                    {sortConfig.column === column ? (sortConfig.direction === "desc" ? "v" : "^") : "<>"}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {pagedRows.map((row, index) => (
                <tr key={index}>
                  {columns.map((column) => (
                    <td
                      key={`${index}-${column}`}
                      className={`${column === "SO" ? "stickyCol1" : ""} ${column === "ASM" ? "stickyCol2" : ""} ${
                        column === "DB Name" || column === "Outlet Name" ? "wrapTextCell" : ""
                      } ${
                        column === "Growth %" ? "noWrapCell" : ""
                      } ${
                        column === "Ach" && parseNumber(row[column]) === 0
                          ? "achZero"
                          : column === "Ach" &&
                              parseNumber(row[column]) > 0 &&
                              normalizeClub(row.Club) === "no club"
                            ? "achNoClub"
                            : ""
                      }`}
                    >
                      {column === "LYSM Tgt" || column === "Ach" ? (
                        formatRounded(row[column])
                      ) : column === "LPPC" ? (
                        getRowLppc(row).toFixed(2)
                      ) : column === "Club" ? (
                        <span className={`clubBadge ${normalizeClub(row[column]).replace(/\s+/g, "-")}`}>
                          {row[column] === null || row[column] === undefined ? "" : String(row[column])}
                        </span>
                      ) : column === "Growth %" ? (
                        <div className="progressCell">
                          <div className="progressTrack">
                            <div className="progressFill" style={{ width: `${parsePercent(row[column])}%` }} />
                          </div>
                          <span className="progressPct">{Math.round(parsePercent(row[column]))}%</span>
                        </div>
                      ) : row[column] === null || row[column] === undefined ? (
                        ""
                      ) : (
                        String(row[column])
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          <div className="pagination">
            <button
              type="button"
              className="pageBtn"
              disabled={safeCurrentPage === 1}
              onClick={() => setCurrentPage(Math.max(1, safeCurrentPage - 1))}
            >
              Prev
            </button>
            {pageList.map((item, idx) =>
              item === "..." ? (
                <span key={`ellipsis-${idx}`} className="pageDots">
                  ...
                </span>
              ) : (
              <button
                key={item}
                type="button"
                className={`pageBtn ${item === safeCurrentPage ? "active" : ""}`}
                onClick={() => setCurrentPage(item)}
              >
                {item}
              </button>
              ),
            )}
            <button
              type="button"
              className="pageBtn"
              onClick={() => setCurrentPage(totalPages)}
            >
              Last
            </button>
            <button
              type="button"
              className="pageBtn"
              disabled={safeCurrentPage === totalPages}
              onClick={() => setCurrentPage(Math.min(totalPages, safeCurrentPage + 1))}
            >
              Next
            </button>
          </div>
        </div>
      </div>
    </main>
  );
}
