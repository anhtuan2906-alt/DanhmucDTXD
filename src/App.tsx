import React, { useState, useEffect, useRef, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx-js-style';
import { Settings, RefreshCw, FilePlus, Database, Download, ExternalLink, AlertCircle, X, ChevronDown, Eye, Lock, User as UserIcon, LogIn, LogOut, Filter, MoreVertical, Loader2, Plus } from 'lucide-react';
import { EVN_HCMC_LOGO } from "./assets/logo";

const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/export?format=csv&gid=0';
const PROJECTS_CSV_URL = 'https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/export?format=csv&gid=1152018861'; // Using gid for 'Thông tin theo MCT' if known, or gviz. Let's use gviz to be safe.
const PROJECTS_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent('Thông tin theo MCT');
const USERS_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent('user');

const PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(SHEET_CSV_URL)}`;
const PROJECTS_PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(PROJECTS_GVIZ_URL)}`;
const USERS_PROXY_URL = `https://corsproxy.io/?${encodeURIComponent(USERS_GVIZ_URL)}`;

const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx7ZFvn9_dAH1AXlz0ZRao-Q65EC6E90pCe5FNnCDODCJlNeVtHRJk_zoV1x8-ZutM0YA/exec';

interface User {
  FullName: string;
  username: string;
  AllowedProjects: string;
}

export default function App() {
  const [currentUser, setCurrentUser] = useState<User | null>(() => {
    const savedUser = localStorage.getItem('currentUser');
    return savedUser ? JSON.parse(savedUser) : null;
  });
  const [users, setUsers] = useState<any[]>([]);
  const [loginForm, setLoginForm] = useState({ username: '', password: '' });
  const [loginError, setLoginError] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  const [data, setData] = useState<string[][]>([]);
  const [projectInfo, setProjectInfo] = useState<Record<string, string>>({});
  const [availableProjects, setAvailableProjects] = useState<Record<string, string>[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [scriptUrl, setScriptUrl] = useState(localStorage.getItem('appsScriptUrl') || DEFAULT_SCRIPT_URL);
  const [showSettings, setShowSettings] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [actionLoading, setActionLoading] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [projectSearchTerm, setProjectSearchTerm] = useState('');
  const [lastUpdated, setLastUpdated] = useState<string>('');
  const [collapsedSections, setCollapsedSections] = useState<Set<string>>(new Set());
  const [selectedSection, setSelectedSection] = useState<string>('all');
  const [showMobileFilters, setShowMobileFilters] = useState(false);

  const [showDropdown, setShowDropdown] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const [showImportProjectDropdown, setShowImportProjectDropdown] = useState(false);
  const [importProjectSearchTerm, setImportProjectSearchTerm] = useState('');
  const importProjectDropdownRef = useRef<HTMLDivElement>(null);

  const [showImportPBDropdown, setShowImportPBDropdown] = useState(false);
  const [importPBSearchTerm, setImportPBSearchTerm] = useState('');
  const importPBDropdownRef = useRef<HTMLDivElement>(null);

  const [showImportNDDropdown, setShowImportNDDropdown] = useState(false);
  const [importNDSearchTerm, setImportNDSearchTerm] = useState('');
  const importNDDropdownRef = useRef<HTMLDivElement>(null);

  const [showAddPBDropdown, setShowAddPBDropdown] = useState(false);
  const [addPBSearchTerm, setAddPBSearchTerm] = useState('');
  const addPBDropdownRef = useRef<HTMLDivElement>(null);
  const dropdownMenuRef = useRef<HTMLDivElement>(null);
  const [activeDropdownMenu, setActiveDropdownMenu] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const clickCountRef = useRef(0);
  const lastClickTimeRef = useRef(0);
  const projectNameRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (projectNameRef.current) {
      projectNameRef.current.style.height = 'auto';
      projectNameRef.current.style.height = `${projectNameRef.current.scrollHeight}px`;
    }
  }, [projectInfo["Tên dự án/công trình"], projectSearchTerm, showDropdown]);

  const handleLogoClick = () => {
    const now = Date.now();
    if (now - lastClickTimeRef.current < 500) {
      clickCountRef.current += 1;
    } else {
      clickCountRef.current = 1;
    }
    lastClickTimeRef.current = now;

    if (clickCountRef.current >= 5) {
      setShowSettings(prev => !prev);
      clickCountRef.current = 0;
    }
  };

  useEffect(() => {
    fetchUsers();
    
    // Handle click outside for custom dropdown
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowDropdown(false);
      }
      if (importProjectDropdownRef.current && !importProjectDropdownRef.current.contains(event.target as Node)) {
        setShowImportProjectDropdown(false);
      }
      if (importPBDropdownRef.current && !importPBDropdownRef.current.contains(event.target as Node)) {
        setShowImportPBDropdown(false);
      }
      if (importNDDropdownRef.current && !importNDDropdownRef.current.contains(event.target as Node)) {
        setShowImportNDDropdown(false);
      }
      if (addPBDropdownRef.current && !addPBDropdownRef.current.contains(event.target as Node)) {
        setShowAddPBDropdown(false);
      }
      if (dropdownMenuRef.current && !dropdownMenuRef.current.contains(event.target as Node)) {
        setActiveDropdownMenu(null);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  useEffect(() => {
    if (currentUser) {
      fetchData();
    }
  }, [currentUser]);

  const fetchUsers = async () => {
    try {
      const timestamp = new Date().getTime();
      const urlWithCacheBuster = `${USERS_GVIZ_URL}&t=${timestamp}`;
      let response;
      try {
        response = await fetch(urlWithCacheBuster);
        if (!response.ok) throw new Error('Direct fetch failed');
      } catch (e) {
        try {
          response = await fetch(`${USERS_PROXY_URL}&t=${timestamp}`);
          if (!response.ok) throw new Error('Proxy fetch failed');
        } catch (e2) {
          response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(urlWithCacheBuster)}`);
        }
      }
      if (response && response.ok) {
        const csvText = await response.text();
        Papa.parse(csvText, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            setUsers(results.data);
          }
        });
      }
    } catch (err) {
      console.error("Failed to fetch users:", err);
    }
  };

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoggingIn(true);
    setLoginError(null);

    // Simulate a small delay for better UX
    setTimeout(() => {
      const user = users.find(u => u.username === loginForm.username && u.Password === loginForm.password);
      if (user) {
        const userData = {
          FullName: user.FullName,
          username: user.username,
          AllowedProjects: user.AllowedProjects || ''
        };
        setCurrentUser(userData);
        localStorage.setItem('currentUser', JSON.stringify(userData));
        setLoginForm({ username: '', password: '' });
      } else {
        setLoginError('Tên đăng nhập hoặc mật khẩu không chính xác.');
      }
      setIsLoggingIn(false);
    }, 800);
  };

  const [importForm, setImportForm] = useState({
    phongBan: 'QLĐT',
    tenCongTrinh: '',
    maCongTrinh: '',
    noiDung: '',
    soHieu: '',
    ngay: new Date().toISOString().split('T')[0]
  });
  const [importFile, setImportFile] = useState<File | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [importMessage, setImportMessage] = useState<{type: 'error' | 'success', text: string} | null>(null);
  // Centralized Sheets Data States
  const [danhMucMauData, setDanhMucMauData] = useState<any[]>([]);
  const [danhMucBoSungData, setDanhMucBoSungData] = useState<any[]>([]);
  const [dataHoSoData, setDataHoSoData] = useState<any[]>([]);

  const [showAddContentModal, setShowAddContentModal] = useState(false);
  const [addContentForm, setAddContentForm] = useState({
    giaiDoan: '',
    dauMuc: '',
    noiDung: '',
    phongBan: 'QLĐT',
    positionMode: 'end',
    targetItem: ''
  });
  const [isAddingContent, setIsAddingContent] = useState(false);
  const [addContentMessage, setAddContentMessage] = useState<{type: 'error' | 'success', text: string} | null>(null);

  // Generic CSV Fetcher
  const fetchGoogleSheetCsv = async (sheetName: string): Promise<any[]> => {
    const timestamp = new Date().getTime();
    const url = `https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}&t=${timestamp}`;
    
    let response;
    try {
      response = await fetch(url);
      if (!response.ok) throw new Error('Direct fetch failed');
    } catch (e) {
      try {
        response = await fetch(`https://corsproxy.io/?${encodeURIComponent(url)}`);
        if (!response.ok) throw new Error('Proxy fetch failed');
      } catch (e2) {
        try {
          response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`);
        } catch (e3) {
          response = await fetch(`https://thingproxy.freeboard.io/fetch/${encodeURIComponent(url)}`);
        }
      }
    }

    if (!response || !response.ok) return [];
    
    const text = await response.text();
    return new Promise((resolve) => {
      Papa.parse(text, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          resolve(results.data as any[]);
        }
      });
    });
  };

  const fetchHoSoLinks = async (): Promise<Map<string, string>> => {
    const timestamp = new Date().getTime();
    const url = `https://docs.google.com/spreadsheets/d/1B237SBdWeaQvc0GWH7hwcJI9ztiSxdBxXFbN4nBnxzU/htmlview/sheet?headers=true&gid=35605870&t=${timestamp}`;
    
    let html = '';
    let response;
    try {
      response = await fetch(url);
      if (!response.ok) throw new Error('Direct fetch failed');
      html = await response.text();
    } catch (e) {
      try {
        response = await fetch(`https://corsproxy.io/?${encodeURIComponent(url)}`);
        if (!response.ok) throw new Error('Proxy fetch failed');
        html = await response.text();
      } catch (e2) {
        try {
          response = await fetch(`https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`);
          if (!response.ok) throw new Error();
          html = await response.text();
        } catch (e3) {
          try {
            response = await fetch(`https://thingproxy.freeboard.io/fetch/${encodeURIComponent(url)}`);
            html = await response.text();
          } catch(e4) {
             return new Map();
          }
        }
      }
    }

    if (!html) return new Map();

    const map = new Map<string, string>();
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const rows = doc.querySelectorAll('table tbody tr');
      
      rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length > 5) {
          const maCT = cells[0].textContent?.replace(/\s+/g, ' ').trim() || '';
          const noiDung = cells[2].textContent?.replace(/\s+/g, ' ').trim() || '';
          const soVB = cells[3].textContent?.replace(/\s+/g, ' ').trim() || '';
          const key = `${maCT}|${noiDung}|${soVB}`;
          
          const linkTag = cells[5].querySelector('a');
          if (linkTag && linkTag.href) {
            let href = linkTag.href;
            if (href.includes('google.com/url?')) {
              try {
                const qs = href.split('?')[1];
                const params = new URLSearchParams(qs);
                const q = params.get('q');
                if (q) href = q;
              } catch(e) {}
            }
            map.set(key, href);
          }
        }
      });
    } catch(e) {
      console.warn("Failed to parse htmlview", e);
    }
    return map;
  };

  const [giaiDoanData, setGiaiDoanData] = useState<any[]>([]);
  const [dauMucData, setDauMucData] = useState<any[]>([]);

  const fetchUnifiedData = async () => {
    setLoading(true);
    try {
      const [mauObj, boSungObj, hoSoObj, gdObj, dmObj, hoSoLinks] = await Promise.all([
        fetchGoogleSheetCsv('Danh_Muc_Mau'),
        fetchGoogleSheetCsv('Danh_Muc_Bo_Sung'),
        fetchGoogleSheetCsv('Data_Ho_So'),
        fetchGoogleSheetCsv('Giaidoan'),
        fetchGoogleSheetCsv('Daumuc'),
        fetchHoSoLinks()
      ]);
      
      hoSoObj.forEach((row, idx) => {
        if (idx === 0) return;
        const maCT = (row[0] || '').toString().replace(/\s+/g, ' ').trim();
        const noiDung = (row[2] || '').toString().replace(/\s+/g, ' ').trim();
        const soVB = (row[3] || '').toString().replace(/\s+/g, ' ').trim();
        const key = `${maCT}|${noiDung}|${soVB}`;
        if (hoSoLinks.has(key)) {
            row[5] = hoSoLinks.get(key);
        }
      });
      
      setDanhMucMauData(mauObj);
      setDanhMucBoSungData(boSungObj);
      setDataHoSoData(hoSoObj);
      setGiaiDoanData(gdObj);
      setDauMucData(dmObj);
    } catch (e) {
      console.warn("Could not fetch some data sheets", e);
    } finally {
      setLoading(false);
      setLastUpdated(new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
    }
  };

  // Run on mount or when user logs in
  useEffect(() => {
    if (currentUser) {
      fetchUnifiedData();
      fetchProjects();
    }
  }, [currentUser]);

  // Derived available options for dropdowns based on new architecture
  const rawDropdownData = useMemo(() => {
    const valid: {phongBan: string, noiDung: string}[] = [];
    
    // Parse Danh_Muc_Mau
    danhMucMauData.forEach((row, idx) => {
      if (idx === 0) return; // skip header
      const noiDung = row[1]?.trim();
      const phongBan = row[2]?.trim();
      if (phongBan && noiDung) valid.push({ phongBan, noiDung });
    });

    // Parse Danh_Muc_Bo_Sung 
    danhMucBoSungData.forEach((row, idx) => {
      if (idx === 0) return; // skip header
      const maCT = row[0]?.trim();
      if (importForm.maCongTrinh && maCT !== importForm.maCongTrinh) return;
      let noiDung = row[3]?.trim();
      
      if (noiDung) {
         const sortMatch = noiDung.match(/(.*?)\|SORTED\|(before|after)\|(.*)/);
         if (sortMatch) noiDung = sortMatch[1];
      }
      
      const phongBan = row[4]?.trim();
      if (phongBan && noiDung) valid.push({ phongBan, noiDung });
    });

    return valid;
  }, [danhMucMauData, danhMucBoSungData, importForm.maCongTrinh]);

  // Pre-fill the import form when modal opens
  useEffect(() => {
    if (showImportModal) {
      setImportForm(prev => ({
        ...prev,
        tenCongTrinh: projectInfo["Tên dự án/công trình"] || '',
        maCongTrinh: projectInfo["Mã công trình"] || ''
      }));
    }
  }, [showImportModal, projectInfo]);

  const handleImportSubmit = async () => {
    if (!scriptUrl) {
      setImportMessage({ type: 'error', text: 'Vui lòng cấu hình Apps Script Web App URL.' });
      return;
    }
    
    if (!importForm.maCongTrinh || !importForm.phongBan || !importForm.noiDung) {
      setImportMessage({ type: 'error', text: 'Vui lòng điền mã công trình, phòng ban và nội dung.' });
      return;
    }

    if (dataHoSoData.length > 0) {
      const hoSoHeaders = dataHoSoData[0] || [];
      const getIdx = (headers: string[], name: string, defaultIdx: number) => {
        const lowerName = name.toLowerCase();
        let idx = headers.findIndex((h: string) => h && h.toString().toLowerCase().trim() === lowerName);
        if (idx === -1) {
          idx = headers.findIndex((h: string) => h && h.toString().toLowerCase().trim().includes(lowerName));
        }
        return idx !== -1 ? idx : defaultIdx;
      };
      const idxMaCT = getIdx(hoSoHeaders, 'mã ct', 0);
      const idxPB = getIdx(hoSoHeaders, 'phòng ban', 1);
      const idxND = getIdx(hoSoHeaders, 'nội dung', 2);

      const isDuplicate = dataHoSoData.slice(1).some(row => {
        const rowMaCT = (row[idxMaCT] || '').trim();
        const rowPB = (row[idxPB] || '').trim();
        const rowND = (row[idxND] || '').trim();
        
        return rowMaCT === importForm.maCongTrinh.trim() &&
               rowPB === importForm.phongBan.trim() &&
               rowND === importForm.noiDung.trim();
      });

      if (isDuplicate) {
        setImportMessage({ type: 'error', text: 'Văn bản đã được nhập. Từ chối thực hiện.' });
        return;
      }
    }

    setIsImporting(true);
    setImportMessage(null);

    try {
      let base64File = '';
      let fileName = '';
      let mimeType = '';

      if (importFile) {
        const reader = new FileReader();
        reader.readAsDataURL(importFile);
        await new Promise((resolve, reject) => {
          reader.onload = () => resolve(null);
          reader.onerror = error => reject(error);
        });
        const result = reader.result as string;
        base64File = result.split(',')[1];
        fileName = importFile.name;
        mimeType = importFile.type;
      }

      // Format date to dd/mm/yyyy
      let formattedDate = importForm.ngay;
      if (formattedDate && formattedDate.includes('-')) {
        const [y, m, d] = formattedDate.split('-');
        formattedDate = `${d}/${m}/${y}`;
      }

      // Pass the folder ID and data
      const payload = {
        action: 'importData',
        phongBan: importForm.phongBan,
        projectCode: importForm.maCongTrinh,
        noiDung: importForm.noiDung,
        coQuan: '',
        soHieu: importForm.soHieu,
        ngay: formattedDate,
        fileName: fileName,
        mimeType: mimeType,
        fileData: base64File,
        targetFolderId: '1gK37pNEA0G4--iocJNyiUMgwrPAxDqJz'
      };

      const startTime = Date.now();
      try {
        await fetch(scriptUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain;charset=utf-8',
          },
          body: JSON.stringify(payload),
          mode: 'no-cors'
        });
      } catch (err: any) {
        // When using no-cors with Google Apps Script, successful executions often end in a redirect 
        // that gets blocked by the browser (CORS, Adblockers, etc.), throwing 'Failed to fetch'.
        // If the request took more than 1 second, it almost certainly reached the server and executed.
        if (err.message === 'Failed to fetch' && Date.now() - startTime > 1000) {
          console.warn('Apps Script execution likely succeeded despite Failed to fetch error.');
        } else {
          throw err;
        }
      }

      // Show success, and schedule a data fetch
      setImportMessage({ type: 'success', text: 'Lưu dữ liệu thành công!' });
      
      // Reset form variables to original states while preserving the active department and project
      setImportForm(prev => ({
        ...prev, 
        soHieu: '', 
        noiDung: '', 
        ngay: new Date().toISOString().split('T')[0]
      }));
      setImportFile(null);
      setImportProjectSearchTerm('');
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
      
      // Polling for update (simplified, just fetch once)
      setTimeout(() => fetchData(), 1500);
      
    } catch (err: any) {
      console.error(err);
      setImportMessage({ type: 'error', text: `Lỗi: ${err.message}. (Nếu dữ liệu đã lên Drive, có thể bỏ qua lỗi này)` });
    } finally {
      setIsImporting(false);
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem('currentUser');
    setProjectInfo({});
    setData([]);
  };

  const fetchData = async () => {
    // Legacy polling function used by import modal. Now simply refreshes the Data_Ho_So.
    try {
      const [hoSoObj, hoSoLinks] = await Promise.all([
        fetchGoogleSheetCsv('Data_Ho_So'),
        fetchHoSoLinks()
      ]);
      hoSoObj.forEach((row, idx) => {
        if (idx === 0) return;
        const maCT = (row[0] || '').toString().replace(/\s+/g, ' ').trim();
        const noiDung = (row[2] || '').toString().replace(/\s+/g, ' ').trim();
        const soVB = (row[3] || '').toString().replace(/\s+/g, ' ').trim();
        const key = `${maCT}|${noiDung}|${soVB}`;
        if (hoSoLinks.has(key)) {
            row[5] = hoSoLinks.get(key);
        }
      });
      setDataHoSoData(hoSoObj);
    } catch (e) {
      console.warn("Failed to update Data_Ho_So", e);
    }
  };

  const fetchProjects = async () => {
    try {
      const rows = await fetchGoogleSheetCsv('Thongtinduan');
      const projects: Record<string, string>[] = [];
      
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const projectCode = row[1]?.trim() || "";
        
        if (currentUser) {
          const allowedProjects = currentUser.AllowedProjects.split(';').map((p: string) => p.trim()).filter((p: string) => p !== '');
          if (allowedProjects.length > 0 && !allowedProjects.includes(projectCode)) continue;
        }

        if (row && row.length >= 3 && row[2] && row[2].trim() !== "" && row[2] !== "Tên công trình") {
          projects.push({
            "Tên dự án/công trình": row[2].trim(),
            "Mã công trình": projectCode,
            "Chủ Đầu Tư": row[4]?.trim() || "",
            "Địa điểm xây dựng": row[3]?.trim() || "",
            "Đơn vị TV Thiết Kế": row[5]?.trim() || "",
            "Đơn vị TV Giám sát": row[6]?.trim() || "",
            "Đơn vị thi công": row[7]?.trim() || ""
          });
        }
      }
      
      const uniqueProjects = projects.filter((v, i, a) => 
        a.findIndex(t => t["Mã công trình"] === v["Mã công trình"]) === i
      );
      
      setAvailableProjects(uniqueProjects);

      setProjectInfo(prev => {
        if (!prev["Mã công trình"] && uniqueProjects.length > 0) {
          return uniqueProjects[0];
        }
        return prev;
      });
    } catch (err) {
      console.error("Failed to fetch projects list:", err);
    }
  };

  // Local Assembly Engine: Joins Templates and Data_Ho_So instantly
  useEffect(() => {
    if (!projectInfo["Mã công trình"]) {
      setData([]);
      return;
    }

    const selectedMaCT = projectInfo["Mã công trình"];

    // Extract mappings
    type GD = { ma: string, name: string };
    type DM = { ma: string, name: string };

    const gdList: GD[] = giaiDoanData.slice(1).map(row => ({
      ma: row[0]?.trim() || '',
      name: row[1]?.trim() || ''
    })).filter(x => x.ma);

    const dmList: DM[] = dauMucData.slice(1).map(row => ({
      ma: row[0]?.trim() || '',
      name: row[1]?.trim() || ''
    })).filter(x => x.ma);

    const findGD = (maNoiDung: string) => {
      let match = null;
      for (const gd of gdList) {
        if (maNoiDung.startsWith(gd.ma)) {
          if (!match || gd.ma.length > match.ma.length) match = gd;
        }
      }
      return match || { ma: 'Z_Khac', name: 'Giai đoạn chưa phân loại' };
    };

    const findDM = (maNoiDung: string) => {
      let match = null;
      for (const dm of dmList) {
        if (maNoiDung.startsWith(dm.ma)) {
          if (!match || dm.ma.length > match.ma.length) match = dm;
        }
      }
      return match || { ma: 'Z_Khac', name: 'Đầu mục chưa phân loại' };
    };

    const resolveGD = (val: string) => {
      const match = gdList.find(x => x.ma === val || x.name === val);
      return match || findGD(val);
    };

    const resolveDM = (val: string) => {
      const match = dmList.find(x => x.ma === val || x.name === val);
      return match || findDM(val);
    };

    const getIdx = (headers: string[], name: string, defaultIdx: number) => {
      const lowerName = name.toLowerCase();
      let idx = headers.findIndex((h: string) => h && h.toString().toLowerCase().trim() === lowerName);
      if (idx === -1) {
        idx = headers.findIndex((h: string) => {
          if (!h) return false;
          const hStr = h.toString().toLowerCase().trim();
          if (hStr.includes(lowerName)) {
            if (!lowerName.includes('mã') && hStr.includes('mã')) return false;
            return true;
          }
          return false;
        });
      }
      return idx >= 0 ? idx : defaultIdx;
    };
    
    // 1. Core Framework
    const mauHeaders = danhMucMauData[0] || [];
    const idxMauMa = getIdx(mauHeaders, 'mã nội dung', 0);
    const idxMauND = getIdx(mauHeaders, 'nội dung', 1);
    const idxMauPB = getIdx(mauHeaders, 'phòng ban', 2);

    const baseItems = danhMucMauData.slice(1).map(row => {
      const ma = row[idxMauMa]?.trim() || '';
      const gd = findGD(ma);
      const dm = findDM(ma);
      return {
        gdMa: gd.ma, gdName: gd.name,
        dmMa: dm.ma, dmName: dm.name,
        noiDung: row[idxMauND]?.trim() || '',
        phongBan: row[idxMauPB]?.trim() || ''
      };
    }).filter(x => x.noiDung);

    // 2. Extra Project-specific items
    const boSungHeaders = danhMucBoSungData[0] || [];
    const idxBsMaCT = getIdx(boSungHeaders, 'mã ct', 0);
    const idxBsGD = getIdx(boSungHeaders, 'giai đoạn', 1);
    const idxBsDM = getIdx(boSungHeaders, 'đầu mục', 2);
    const idxBsND = getIdx(boSungHeaders, 'nội dung', 3);
    const idxBsPB = getIdx(boSungHeaders, 'phòng ban', 4);

    const extraItems = danhMucBoSungData.slice(1).map(row => {
      const rawGD = row[idxBsGD]?.trim() || '';
      const rawDM = row[idxBsDM]?.trim() || '';
      const gd = resolveGD(rawGD);
      const dm = resolveDM(rawDM);
      
      let rawNoiDung = row[idxBsND]?.trim() || '';
      let noiDung = rawNoiDung;
      let sortMode = '';
      let sortTarget = '';
      
      const sortMatch = rawNoiDung.match(/(.*?)\|SORTED\|(before|after)\|(.*)/);
      if (sortMatch) {
         noiDung = sortMatch[1];
         sortMode = sortMatch[2];
         sortTarget = sortMatch[3];
      }
      
      return {
        maCT: row[idxBsMaCT]?.trim() || '',
        gdMa: gd.ma, gdName: gd.name,
        dmMa: dm.ma, dmName: dm.name,
        noiDung: noiDung,
        phongBan: row[idxBsPB]?.trim() || '',
        sortMode, sortTarget
      };
    }).filter(x => x.maCT === selectedMaCT && x.noiDung);

    // Grouping structure by Mã
    type ItemType = { noiDung: string, phongBan: string };
    const itemsByGD: Record<string, { name: string, dms: Record<string, { name: string, items: ItemType[] }> }> = {};

    baseItems.forEach(item => {
      if (!itemsByGD[item.gdMa]) itemsByGD[item.gdMa] = { name: item.gdName, dms: {} };
      if (!itemsByGD[item.gdMa].dms[item.dmMa]) itemsByGD[item.gdMa].dms[item.dmMa] = { name: item.dmName, items: [] };
      const exists = itemsByGD[item.gdMa].dms[item.dmMa].items.find(x => x.noiDung === item.noiDung);
      if (!exists) {
        itemsByGD[item.gdMa].dms[item.dmMa].items.push({ noiDung: item.noiDung, phongBan: item.phongBan });
      }
    });
    
    extraItems.forEach(item => {
      if (!itemsByGD[item.gdMa]) itemsByGD[item.gdMa] = { name: item.gdName, dms: {} };
      if (!itemsByGD[item.gdMa].dms[item.dmMa]) itemsByGD[item.gdMa].dms[item.dmMa] = { name: item.dmName, items: [] };
      
      const targetArray = itemsByGD[item.gdMa].dms[item.dmMa].items;
      const exists = targetArray.find(x => x.noiDung === item.noiDung);
      
      if (!exists) {
        let inserted = false;
        if (item.sortMode && item.sortTarget) {
           const targetIdx = targetArray.findIndex(x => x.noiDung === item.sortTarget);
           if (targetIdx !== -1) {
              if (item.sortMode === 'before') {
                 targetArray.splice(targetIdx, 0, { noiDung: item.noiDung, phongBan: item.phongBan });
              } else {
                 targetArray.splice(targetIdx + 1, 0, { noiDung: item.noiDung, phongBan: item.phongBan });
              }
              inserted = true;
           }
        }
        if (!inserted) {
           targetArray.push({ noiDung: item.noiDung, phongBan: item.phongBan });
        }
      }
    });

    // 3. Match Records
    const hoSoHeaders = dataHoSoData[0] || [];
    const idxMaCT = getIdx(hoSoHeaders, 'mã ct', 0);
    const idxPB = getIdx(hoSoHeaders, 'phòng ban', 1);
    const idxND = getIdx(hoSoHeaders, 'nội dung', 2);
    const idxSoVB = getIdx(hoSoHeaders, 'số vb', 3);
    const idxNgayVB = getIdx(hoSoHeaders, 'ngày vb', 4);
    const idxFile = getIdx(hoSoHeaders, 'link', 5);
    const idxNgayTaiLen = getIdx(hoSoHeaders, 'tải lên', 6);
    const idxCoQuan = getIdx(hoSoHeaders, 'cơ quan', 7);

    const records = dataHoSoData.slice(1).map(row => ({
      maCT: row[idxMaCT]?.trim() || '',
      phongBan: row[idxPB]?.trim() || '',
      noiDung: row[idxND]?.trim() || '',
      soVB: row[idxSoVB]?.trim() || '',
      ngayVB: row[idxNgayVB]?.trim() || '',
      fileVB: row[idxFile]?.trim() || '',
      ngayTaiLen: row[idxNgayTaiLen]?.trim() || '',
      coQuan: row[idxCoQuan]?.trim() || ''
    })).filter(x => x.maCT === selectedMaCT);

    // Any records uploaded outside template boundaries
    records.forEach(r => {
      // Find where this noiDung exists in our itemsByGD
      let found = false;
      for (const gd in itemsByGD) {
        for (const dm in itemsByGD[gd].dms) {
          if (itemsByGD[gd].dms[dm].items.some(x => x.noiDung === r.noiDung)) {
            found = true;
          }
        }
      }
      if (!found) {
        // Just put in Khác / Khác
        if (!itemsByGD['Z_Khac']) itemsByGD['Z_Khac'] = { name: 'Giai đoạn khác', dms: {} };
        if (!itemsByGD['Z_Khac'].dms['Z_Khac']) itemsByGD['Z_Khac'].dms['Z_Khac'] = { name: 'Hạng mục khác', items: [] };
        itemsByGD['Z_Khac'].dms['Z_Khac'].items.push({ noiDung: r.noiDung, phongBan: r.phongBan });
      }
    });

    const finalRows: string[][] = [];
    let stt = 1;

    // Use predefined sorting if known, else alphabetical
    const sortedGdMas = Object.keys(itemsByGD).sort((a,b) => a.localeCompare(b));
    
    sortedGdMas.forEach(gdMa => {
      // Create Giai Đoạn Header Row
      const gdName = itemsByGD[gdMa].name;
      finalRows.push(["", gdName.toUpperCase(), `TYPE:GD|${gdMa}`, "", "", "", ""]);

      const sortedDmMas = Object.keys(itemsByGD[gdMa].dms).sort((a,b) => a.localeCompare(b));
      
      sortedDmMas.forEach(dmMa => {
        // Create Đầu mục Header Row
        const dmName = itemsByGD[gdMa].dms[dmMa].name;
        finalRows.push(["", dmName, `TYPE:DM|${gdMa}|${dmMa}`, "", "", "", ""]);

        itemsByGD[gdMa].dms[dmMa].items.forEach(item => {
          const matches = records.filter(r => r.noiDung === item.noiDung);
          
          if (matches.length > 0) {
            matches.forEach(r => {
              // Ensure any link correctly parses
              let finalLink = r.fileVB;
              if (finalLink.includes('HYPERLINK(')) {
                // try to extract URL if using formula
                const urlMatch = finalLink.match(/HYPERLINK\("([^"]+)"/);
                if (urlMatch) finalLink = urlMatch[1];
              }

              finalRows.push([
                stt.toString(), item.noiDung, `TYPE:ITEM|${gdMa}|${dmMa}`, r.soVB, r.ngayVB, finalLink, item.phongBan || r.coQuan || ''
              ]);
              stt++;
            });
          } else {
            finalRows.push([ stt.toString(), item.noiDung, `TYPE:ITEM|${gdMa}|${dmMa}`, "", "", "", item.phongBan || "" ]);
            stt++;
          }
        });
      });
    });

    setData(finalRows);
  }, [projectInfo, danhMucMauData, danhMucBoSungData, dataHoSoData, giaiDoanData, dauMucData]);

  const saveScriptUrl = (url: string) => {
    setScriptUrl(url);
    localStorage.setItem('appsScriptUrl', url);
  };

  // Clean up data for display
  const allRows = data.filter(row => row && row.length > 0 && !!row[2]);

  // Identify sections and assign rows to them
  const sections: string[] = [];
  const rowToGdName = new Map<number, string>();
  let currentGdName = "";

  allRows.forEach((row, index) => {
    if (row[2]?.startsWith('TYPE:GD|')) {
      currentGdName = row[1];
      if (!sections.includes(currentGdName)) {
        sections.push(currentGdName);
      }
    }
    rowToGdName.set(index, currentGdName);
  });

  const rows = allRows.filter((row, index) => {
    let matchesSearch = true;
    if (searchTerm) {
      matchesSearch = row.some(cell => typeof cell === 'string' && cell.toLowerCase().includes(searchTerm.toLowerCase()));
    }
    
    const gdNameOfRow = rowToGdName.get(index) || "";
    const matchesSection = selectedSection === 'all' || gdNameOfRow === selectedSection;
    
    // Always keep GD and DM headers if section matches, letting search filter them loosely
    const isGD = row[2]?.startsWith('TYPE:GD|');
    const isDM = row[2]?.startsWith('TYPE:DM|');
    if ((isGD || isDM) && matchesSection && !searchTerm) {
      return true; // Fast pass if no search
    }

    return matchesSearch && matchesSection;
  });

  // Further filter rows to hide those in collapsed sections
  const visibleRows = rows.filter((row) => {
    if (!row[2]) return true;

    const parts = row[2].split('|');
    const type = parts[0]; 
    const gdMa = parts[1];
    const dmMa = parts[2];

    if (type === 'TYPE:GD') {
      return true;
    }

    if (type === 'TYPE:DM') {
      return !collapsedSections.has(`GD|${gdMa}`);
    }

    if (type === 'TYPE:ITEM') {
      return !collapsedSections.has(`GD|${gdMa}`) && !collapsedSections.has(`DM|${gdMa}|${dmMa}`);
    }

    return true;
  });

  const toggleSection = (id: string) => {
    setCollapsedSections(prev => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="max-w-md w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-200">
          <div className="bg-white p-8 text-center border-b border-slate-100">
            <div className="inline-block mb-4 cursor-default select-none">
              <img 
                src={EVN_HCMC_LOGO} 
                alt="EVN HCMC" 
                className="h-24 w-auto object-contain"
                referrerPolicy="no-referrer"
              />
            </div>
            <h1 className="text-blue-900 text-2xl font-bold uppercase tracking-wider">Đăng nhập hệ thống</h1>
            <p className="text-slate-500 mt-2 text-sm font-medium">Danh mục hồ sơ ĐTXD</p>
          </div>
          
          <form onSubmit={handleLogin} className="p-8 space-y-6">
            {loginError && (
              <div className="bg-red-50 border-l-4 border-red-500 p-4 flex gap-3 items-start animate-in fade-in slide-in-from-top-2">
                <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                <p className="text-sm text-red-700">{loginError}</p>
              </div>
            )}
            
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5">Tên đăng nhập</label>
                <div className="relative">
                  <UserIcon className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                  <input 
                    type="text"
                    required
                    value={loginForm.username}
                    onChange={(e) => setLoginForm(prev => ({ ...prev, username: e.target.value }))}
                    className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-slate-800"
                    placeholder="Nhập tên đăng nhập..."
                  />
                </div>
              </div>
              
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5">Mật khẩu</label>
                <div className="relative">
                  <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                  <input 
                    type="password"
                    required
                    value={loginForm.password}
                    onChange={(e) => setLoginForm(prev => ({ ...prev, password: e.target.value }))}
                    className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-slate-800"
                    placeholder="Nhập mật khẩu..."
                  />
                </div>
              </div>
            </div>
            
            <button 
              type="submit"
              disabled={isLoggingIn}
              className="w-full bg-blue-900 text-white py-3.5 rounded-xl font-bold uppercase tracking-widest hover:bg-blue-800 active:scale-[0.98] transition-all shadow-lg shadow-blue-900/20 flex items-center justify-center gap-2 disabled:opacity-70"
            >
              {isLoggingIn ? (
                <RefreshCw className="w-5 h-5 animate-spin" />
              ) : (
                <LogIn className="w-5 h-5" />
              )}
              <span>Đăng nhập</span>
            </button>
          </form>
          
          <div className="px-8 py-4 bg-slate-50 border-t border-slate-100 text-center">
            <p className="text-xs text-slate-400">© 2026 EVN HCMC - Hệ thống quản lý nội bộ</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex items-center cursor-pointer select-none" onClick={handleLogoClick} title="EVN HCMC">
              <img 
                src={EVN_HCMC_LOGO} 
                alt="EVN HCMC" 
                className="h-16 w-auto object-contain"
                referrerPolicy="no-referrer"
              />
            </div>
            <h1 className="text-xl sm:text-2xl font-bold text-blue-900 uppercase tracking-tight">DANH MỤC HỒ SƠ ĐTXD</h1>
          </div>
          <div className="flex items-center gap-4">
            <div className="hidden sm:flex flex-col items-end mr-2">
              <span className="text-sm font-bold text-slate-700">{currentUser.FullName}</span>
              <span className="text-[10px] text-slate-400 uppercase tracking-wider font-medium">@{currentUser.username}</span>
            </div>
            <button 
              onClick={handleLogout}
              className="p-2.5 bg-slate-100 text-slate-600 rounded-full hover:bg-red-50 hover:text-red-600 transition-all border border-slate-200"
              title="Đăng xuất"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-[1600px] mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Import Modal */}
        {showImportModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
            <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg flex flex-col max-h-[90vh]">
              <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-white rounded-t-2xl">
                <div className="flex items-center gap-2">
                  <Database className="w-5 h-5 text-blue-600" />
                  <h3 className="font-bold text-slate-800 uppercase text-lg">Nhập dữ liệu</h3>
                </div>
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    setImportMessage(null);
                    fetchData();
                  }}
                  className="p-2 hover:bg-slate-100 rounded-full text-slate-400 hover:text-slate-600 transition-all"
                  title="Đóng cửa sổ"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
              
              <div className="p-6 overflow-y-auto flex-1 bg-white pb-64">
                <div className="space-y-5">
                  <div className="relative z-[60]" ref={importPBDropdownRef}>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Chọn Phòng ban/Đơn vị:</label>
                    <input
                      type="text"
                      value={showImportPBDropdown ? importPBSearchTerm : importForm.phongBan}
                      onChange={e => setImportPBSearchTerm(e.target.value)}
                      onFocus={() => {
                        setImportPBSearchTerm('');
                        setShowImportPBDropdown(true);
                      }}
                      placeholder="Nhập hoặc chọn phòng ban..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm bg-white cursor-text"
                    />
                    {showImportPBDropdown && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-xl max-h-60 overflow-y-auto z-[60]">
                        {Array.from(new Set<string>(rawDropdownData.map(d => d.phongBan as string)))
                          .filter(pb => pb && (!importPBSearchTerm || pb.toLowerCase().includes(importPBSearchTerm.toLowerCase())))
                          .sort((a,b) => a.localeCompare(b))
                          .map((pb, idx) => (
                          <div
                            key={idx}
                            onClick={() => {
                              setImportForm({...importForm, phongBan: pb, noiDung: ''});
                              setShowImportPBDropdown(false);
                            }}
                            className="px-4 py-2 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 transition-colors text-sm text-slate-700"
                          >
                            {pb}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  <div className="relative z-[50]" ref={importProjectDropdownRef}>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Tên công trình:</label>
                    <input
                      type="text"
                      value={showImportProjectDropdown ? importProjectSearchTerm : importForm.tenCongTrinh}
                      onChange={e => setImportProjectSearchTerm(e.target.value)}
                      onFocus={() => {
                        setImportProjectSearchTerm('');
                        setShowImportProjectDropdown(true);
                      }}
                      placeholder="Nhập tên để tìm kiếm..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                    {showImportProjectDropdown && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-xl max-h-60 overflow-y-auto z-[60]">
                        {availableProjects.filter(p => {
                          const term = importProjectSearchTerm.toLowerCase();
                          return !term || 
                                 p["Tên dự án/công trình"]?.toLowerCase().includes(term) || 
                                 p["Mã công trình"]?.toLowerCase().includes(term);
                        }).map((project, idx) => (
                          <div
                            key={idx}
                            onClick={() => {
                              setImportForm({
                                ...importForm, 
                                tenCongTrinh: project["Tên dự án/công trình"],
                                maCongTrinh: project["Mã công trình"]
                              });
                              setShowImportProjectDropdown(false);
                            }}
                            className="px-4 py-3 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 transition-colors"
                          >
                            <div className="font-medium text-slate-800">{project["Tên dự án/công trình"]}</div>
                            <div className="text-xs text-slate-500 mt-1">Mã CT: {project["Mã công trình"]}</div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Mã công trình:</label>
                    <input
                      type="text"
                      value={importForm.maCongTrinh}
                      disabled
                      placeholder="Mã công trình sẽ tự động điền..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-500 shadow-sm bg-slate-100 cursor-not-allowed"
                    />
                  </div>

                  <div className="relative z-[40]" ref={importNDDropdownRef}>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Nội dung văn bản:</label>
                    <input
                      type="text"
                      value={showImportNDDropdown ? importNDSearchTerm : importForm.noiDung}
                      onChange={e => setImportNDSearchTerm(e.target.value)}
                      onFocus={() => {
                        setImportNDSearchTerm('');
                        setShowImportNDDropdown(true);
                      }}
                      placeholder="Nhập hoặc chọn nội dung văn bản..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                    {showImportNDDropdown && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-xl max-h-60 overflow-y-auto z-[60]">
                        {Array.from(new Set<string>(rawDropdownData.filter(item => item.phongBan === importForm.phongBan).map(item => item.noiDung as string)))
                          .filter(nd => nd && (!importNDSearchTerm || nd.toLowerCase().includes(importNDSearchTerm.toLowerCase())))
                          .sort((a,b) => a.localeCompare(b))
                          .map((nd, idx) => (
                          <div
                            key={idx}
                            onClick={() => {
                              setImportForm({...importForm, noiDung: nd});
                              setShowImportNDDropdown(false);
                            }}
                            className="px-4 py-2 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 transition-colors text-sm text-slate-700 whitespace-normal"
                          >
                            {nd}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  
                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Số hiệu văn bản:</label>
                    <input
                      type="text"
                      value={importForm.soHieu}
                      onChange={e => setImportForm({...importForm, soHieu: e.target.value})}
                      placeholder="Nhập số văn bản..."
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                  </div>

                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Ngày văn bản:</label>
                    <input
                      type="date"
                      value={importForm.ngay}
                      onChange={e => setImportForm({...importForm, ngay: e.target.value})}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 shadow-sm"
                    />
                  </div>

                  <div>
                    <label className="block mb-1.5 text-sm font-bold text-slate-700">Đính kèm văn bản (PDF, Ảnh, Word...):</label>
                    <input
                      type="file"
                      ref={fileInputRef}
                      onChange={e => {
                        if (e.target.files && e.target.files.length > 0) {
                          setImportFile(e.target.files[0]);
                        } else {
                          setImportFile(null);
                        }
                      }}
                      className="w-full px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-slate-700 bg-white shadow-sm file:mr-4 file:py-1 file:px-3 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-slate-100 file:text-slate-700 hover:file:bg-slate-200 cursor-pointer"
                    />
                  </div>

                  <button
                    onClick={handleImportSubmit}
                    disabled={isImporting}
                    className="w-full py-2.5 bg-blue-600 text-white rounded-md hover:bg-blue-700 font-bold transition-colors disabled:bg-blue-400 mt-2 shadow-sm"
                  >
                    {isImporting ? 'ĐANG GỬI...' : 'GỬI DỮ LIỆU'}
                  </button>

                  {importMessage && (
                    <div className={`mt-4 text-sm font-bold flex items-center justify-center gap-2 ${importMessage.type === 'error' ? 'text-red-600' : 'text-emerald-600'}`}>
                      {importMessage.type === 'error' ? <X className="w-5 h-5 font-bold" /> : <div className="w-2 h-2 rounded-full bg-emerald-600"></div>}
                      {importMessage.text}
                    </div>
                  )}
                </div>
              </div>
              <div className="px-6 py-4 border-t border-slate-200 flex justify-end bg-white rounded-b-2xl">
                <button 
                  onClick={() => {
                    setShowImportModal(false);
                    setImportMessage(null);
                    fetchData();
                  }}
                  className="px-6 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 font-bold transition-colors shadow-sm"
                >
                  HOÀN TẤT
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Add Content Modal */}
        {showAddContentModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4 animate-in fade-in duration-200">
            <div className="bg-white/95 backdrop-blur-xl rounded-2xl shadow-2xl border border-white/50 w-full max-w-sm flex flex-col max-h-[90vh] animate-in zoom-in-95 duration-200">
              <div className="px-5 py-3.5 border-b border-slate-100 flex justify-between items-center rounded-t-2xl bg-white/50">
                <div className="flex items-center gap-2">
                  <div className="p-1.5 bg-emerald-100 rounded-lg text-emerald-600">
                    <FilePlus className="w-4 h-4" />
                  </div>
                  <h3 className="font-bold tracking-tight text-slate-800 text-[15px]">Thêm nội dung</h3>
                </div>
                <button 
                  onClick={() => {
                    setShowAddContentModal(false);
                    setAddContentMessage(null);
                  }}
                  className="p-1.5 hover:bg-slate-100/80 rounded-full text-slate-400 hover:text-slate-600 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
              
              <div className="p-5 overflow-y-auto flex-1 bg-white space-y-4 pb-20">
                <div>
                  <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Tên công trình</label>
                  <input
                    type="text"
                    value={projectInfo["Tên dự án/công trình"] || ''}
                    disabled
                    className="w-full px-3 py-2 border border-slate-100 rounded-lg bg-slate-50 text-slate-600/80 text-sm font-medium opacity-80 cursor-not-allowed"
                  />
                </div>
                <div>
                  <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Giai đoạn</label>
                  <select
                    value={addContentForm.giaiDoan}
                    onChange={e => setAddContentForm({...addContentForm, giaiDoan: e.target.value})}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 text-sm font-medium text-slate-700 bg-white shadow-sm transition-all"
                  >
                    <option value="">-- Chọn Giai đoạn --</option>
                    {giaiDoanData.slice(1).map(row => (
                      <option key={row[0]} value={row[0]}>{row[1]}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Hạng mục văn bản</label>
                  <select
                    value={addContentForm.dauMuc}
                    onChange={e => setAddContentForm({...addContentForm, dauMuc: e.target.value})}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 text-sm font-medium text-slate-700 bg-white shadow-sm transition-all"
                  >
                    <option value="">-- Chọn Hạng mục --</option>
                    {dauMucData.slice(1).map(row => (
                      <option key={row[0]} value={row[0]}>{row[1]}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Nội dung văn bản (Thêm mới)</label>
                  <input
                    type="text"
                    value={addContentForm.noiDung}
                    onChange={e => setAddContentForm({...addContentForm, noiDung: e.target.value})}
                    placeholder="Nhập nội dung..."
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 text-sm font-medium text-slate-800 bg-white shadow-sm transition-all placeholder:text-slate-400 placeholder:font-normal"
                  />
                </div>
                <div>
                  <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Vị trí thêm mới</label>
                  <select
                    value={addContentForm.positionMode}
                    onChange={e => setAddContentForm({...addContentForm, positionMode: e.target.value, targetItem: ''})}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 text-sm font-medium text-slate-700 bg-white shadow-sm transition-all"
                  >
                    <option value="end">Thêm vào cuối danh sách</option>
                    <option value="before">Chèn trước một nội dung</option>
                    <option value="after">Chèn sau một nội dung</option>
                  </select>
                </div>
                {addContentForm.positionMode !== 'end' && addContentForm.giaiDoan && addContentForm.dauMuc && (
                  <div>
                    <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Chọn nội dung hiện tại</label>
                    <select
                      value={addContentForm.targetItem}
                      onChange={e => setAddContentForm({...addContentForm, targetItem: e.target.value})}
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 text-sm font-medium text-slate-700 bg-white shadow-sm transition-all"
                    >
                      <option value="">-- Chọn nội dung --</option>
                      {/* Calculate available items for chosen Gd and Dm from data */}
                      {(() => {
                        // find items in this section
                        const availableItems = data.filter(r => r[2] === `TYPE:ITEM|${addContentForm.giaiDoan}|${addContentForm.dauMuc}`);
                        return availableItems.map((r, i) => (
                           <option key={i} value={r[1]}>{r[1]}</option>
                        ));
                      })()}
                    </select>
                  </div>
                )}
                <div className="relative z-[60]" ref={addPBDropdownRef}>
                  <label className="block mb-1 text-[11px] font-bold text-slate-500 uppercase tracking-wider">Phòng ban phụ trách</label>
                  <input
                      type="text"
                      value={showAddPBDropdown ? addPBSearchTerm : addContentForm.phongBan}
                      onChange={e => setAddPBSearchTerm(e.target.value)}
                      onFocus={() => {
                        setAddPBSearchTerm('');
                        setShowAddPBDropdown(true);
                      }}
                      placeholder="Nhập/chọn phòng ban..."
                      className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 text-sm font-medium text-slate-800 bg-white shadow-sm transition-all placeholder:text-slate-400 cursor-text"
                    />
                    {showAddPBDropdown && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white/95 backdrop-blur-xl border border-slate-200 rounded-lg shadow-[0_10px_40px_-10px_rgba(0,0,0,0.1)] max-h-48 overflow-y-auto z-[60]">
                        {Array.from(new Set<string>(rawDropdownData.map(d => d.phongBan as string)))
                          .filter(pb => pb && (!addPBSearchTerm || pb.toLowerCase().includes(addPBSearchTerm.toLowerCase())))
                          .sort((a,b) => a.localeCompare(b))
                          .map((pb, idx) => (
                          <div
                            key={idx}
                            onClick={() => {
                              setAddContentForm({...addContentForm, phongBan: pb});
                              setShowAddPBDropdown(false);
                            }}
                            className="px-3 py-2 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 transition-colors text-[13px] font-medium text-slate-700"
                          >
                            {pb}
                          </div>
                        ))}
                      </div>
                    )}
                </div>

                <button
                  onClick={async () => {
                     if (!scriptUrl) {
                        setAddContentMessage({ type: 'error', text: 'Vui lòng cấu hình Apps Script Web App URL.' });
                        return;
                     }
                     if (!projectInfo["Mã công trình"]) {
                        setAddContentMessage({ type: 'error', text: 'Vui lòng chọn công trình trước.' });
                        return;
                     }
                     if (!addContentForm.giaiDoan || !addContentForm.dauMuc || !addContentForm.noiDung) {
                        setAddContentMessage({ type: 'error', text: 'Vui lòng điền đầy đủ Giai đoạn, Hạng mục và Nội dung.' });
                        return;
                     }
                     if (addContentForm.positionMode !== 'end' && !addContentForm.targetItem) {
                        setAddContentMessage({ type: 'error', text: 'Vui lòng chọn nội dung hiện tại để chèn.' });
                        return;
                     }
                     
                     let finalNoiDung = addContentForm.noiDung;
                     if (addContentForm.positionMode !== 'end' && addContentForm.targetItem) {
                        finalNoiDung += `|SORTED|${addContentForm.positionMode}|${addContentForm.targetItem}`;
                     }
                     
                     setIsAddingContent(true);
                     setAddContentMessage(null);
                     try {
                        const payload = {
                          action: 'addCustomCategory',
                          maCongTrinh: projectInfo["Mã công trình"],
                          giaiDoan: addContentForm.giaiDoan,
                          dauMuc: addContentForm.dauMuc,
                          noiDung: finalNoiDung,
                          phongBan: addContentForm.phongBan
                        };
                        
                        const startTime = Date.now();
                        try {
                          await fetch(scriptUrl, {
                            method: 'POST',
                            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                            body: JSON.stringify(payload),
                            mode: 'no-cors'
                          });
                        } catch (err: any) {
                          if (err.message === 'Failed to fetch' && Date.now() - startTime > 1000) {
                            console.warn('Apps Script execution likely succeeded despite Failed to fetch error.');
                          } else {
                            throw err;
                          }
                        }

                        setAddContentMessage({ type: 'success', text: 'Đã hoàn tất!' });
                        setAddContentForm(prev => ({...prev, noiDung: '', targetItem: '', positionMode: 'end'}));
                        
                        // Refetch Data 
                        setTimeout(() => {
                           fetchUnifiedData();
                        }, 1000);

                     } catch(e: any) {
                        setAddContentMessage({ type: 'error', text: 'Gặp lỗi: ' + e.message });
                     } finally {
                        setIsAddingContent(false);
                     }
                  }}
                  disabled={isAddingContent}
                  className="w-full mt-4 flex items-center justify-center gap-2 px-4 py-2 bg-slate-900 text-white rounded-lg hover:bg-slate-800 font-medium text-sm transition-all disabled:opacity-70 disabled:cursor-not-allowed shadow-[0_4px_14px_0_rgba(15,23,42,0.15)] hover:shadow-[0_6px_20px_rgba(15,23,42,0.2)]"
                >
                  {isAddingContent ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin text-white/70" />
                      <span>Đang lưu...</span>
                    </>
                  ) : (
                    <>
                      <span>Thêm nội dung</span>
                    </>
                  )}
                </button>

                {addContentMessage && (
                  <div className={`p-3 rounded-lg text-xs font-medium border ${addContentMessage.type === 'error' ? 'bg-red-50 text-red-700 border-red-100' : 'bg-emerald-50 text-emerald-700 border-emerald-100'}`}>
                    {addContentMessage.text}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* Settings Panel */}
        {showSettings && (
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 mb-8 animate-in slide-in-from-top-4 fade-in duration-200">
            <h2 className="text-lg font-medium mb-4 flex items-center gap-2">
              <Settings className="w-5 h-5 text-slate-500" />
              Cấu hình kết nối Google Apps Script
            </h2>
            <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 mb-4 text-sm text-amber-800 flex gap-3">
              <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
              <div>
                <p className="font-medium mb-1">Tại sao cần cấu hình URL này?</p>
                <p className="mb-2">Để các nút chức năng hoạt động, bạn cần triển khai mã Apps Script trên Google Sheet của bạn dưới dạng <strong>Web App</strong> và dán URL vào đây.</p>
                <ol className="list-decimal ml-5 space-y-1">
                  <li>Mở Google Sheet, vào <strong>Tiện ích mở rộng &gt; Apps Script</strong>.</li>
                  <li>Thêm hàm <code>doGet(e)</code> để xử lý các tham số <code>action</code> ('create', 'import', 'export').</li>
                  <li>Chọn <strong>Triển khai &gt; Triển khai mới</strong>, chọn loại <strong>Ứng dụng web</strong>.</li>
                  <li>Quyền truy cập: <strong>Bất kỳ ai</strong>.</li>
                  <li>Sao chép URL Web App và dán vào ô bên dưới.</li>
                </ol>
              </div>
            </div>
            <div className="flex gap-3">
              <input 
                type="url" 
                value={scriptUrl}
                onChange={(e) => saveScriptUrl(e.target.value)}
                placeholder="https://script.google.com/macros/s/.../exec"
                className="flex-1 px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none transition-shadow"
              />
              <button 
                onClick={() => setShowSettings(false)}
                className="px-4 py-2 bg-slate-800 text-white rounded-lg hover:bg-slate-700 font-medium transition-colors"
              >
                Lưu & Đóng
              </button>
            </div>
          </div>
        )}

        {/* MAIN LAYOUT WRAPPER */}
        <div className="flex flex-col lg:flex-row gap-6 items-stretch relative">
          
          {/* LEFT SIDEBAR (Desktop) */}
          <div className="w-full lg:w-[400px] shrink-0 flex flex-col gap-6 w-full self-stretch">
            {/* Action Buttons */}
            <div className="hidden sm:block bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <div className="grid grid-cols-2 gap-3 w-full">
                <button 
                  onClick={() => {
                    if (!scriptUrl) {
                      alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                      setShowSettings(true);
                      return;
                    }
                    setShowImportModal(true);
                  }}
                  className="flex flex-row items-center justify-center gap-1.5 px-3 py-2 bg-gradient-to-br from-blue-600 to-blue-700 text-white rounded-lg hover:from-blue-700 hover:to-blue-800 font-bold shadow-sm hover:shadow-md active:scale-95 transition-all duration-200 w-full"
                >
                  <FilePlus className="w-4 h-4" />
                  <span className="uppercase tracking-wider text-[11px] sm:text-xs">Nhập DL</span>
                </button>
                
                <button 
                  onClick={() => {
                    try {
                      const wb = XLSX.utils.book_new();
                      const wsData: any[][] = [];
                      const merges: any[] = [];
                      
                      // Thêm 2 hàng đầu tiên
                      wsData.push(['DANH MỤC HỒ SƠ', '', '', '', '', '', '', '']); // Dòng 1
                      wsData.push(['', '', '', '', '', '', '', '']);                 // Dòng 2
                      merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 7 } });         // Merge A1:H1
                      
                      let rowIndex = 2;
                      // Mảng thông tin dự án
                      Object.entries(projectInfo).forEach(([key, value]) => {
                        wsData.push([`${key}:`, '', value, '', '', '', '', '']); // 8 cột
                        merges.push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: 1 } });
                        merges.push({ s: { r: rowIndex, c: 2 }, e: { r: rowIndex, c: 3 } }); // merge C:D
                        rowIndex++;
                      });
                      
                      // Bù thêm các dòng trống nếu thông tin dự án ít hơn để Header chính xác ở dòng 11 (rowIndex 10)
                      while (rowIndex < 9) {
                         wsData.push(['', '', '', '', '', '', '', '']);
                         merges.push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: 1 } });
                         merges.push({ s: { r: rowIndex, c: 2 }, e: { r: rowIndex, c: 3 } }); // merge C:D
                         rowIndex++;
                      }
                      
                      // Dòng 10 (trống) trước khi vào header
                      wsData.push(['', '', '', '', '', '', '', '']);
                      rowIndex++;
                      
                      const headerRowIndex = rowIndex; // = 10 (hiển thị là số 11 trên excel)
                      wsData.push(['STT', 'TÊN VĂN BẢN', '', 'SỐ VB', 'NGÀY VB', 'FILE VB', 'CQBH', 'GHI CHÚ']);
                      merges.push({ s: { r: headerRowIndex, c: 1 }, e: { r: headerRowIndex, c: 2 } });
                      
                      const borderStyle = { style: 'thin', color: { rgb: "000000" } };
                      const borderObj = { top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle };
                      
                      data.forEach((row) => {
                         let fileLink = row[5];
                         let hasLink = fileLink && (fileLink.startsWith('http://') || fileLink.startsWith('https://'));
                         let fileDisp = hasLink ? "Xem File" : (fileLink || "");
                         
                         wsData.push([
                           { v: row[0] || '', t: 's' }, 
                           { v: row[1] || '', t: 's' }, 
                           { v: '', t: 's' }, 
                           { v: row[3] || '', t: 's' }, 
                           { v: row[4] || '', t: 's' }, 
                           { v: fileDisp, t: 's' }, 
                           { v: row[6] || '', t: 's' }, 
                           { v: '', t: 's' }
                         ]);
                      });
                      
                      const ws = XLSX.utils.aoa_to_sheet(wsData);
                      
                      // Gộp ô TÊN VĂN BẢN B10:C...
                      for(let r = headerRowIndex + 1; r < wsData.length; r++) {
                         merges.push({ s: { r: r, c: 1 }, e: { r: r, c: 2 } });
                      }
                      
                      ws['!merges'] = merges;
                      
                      ws['!cols'] = [
                        { wch: 8 },   // A: STT
                        { wch: 15 },  // B: TÊN VĂN BẢN (merge C)
                        { wch: 50 },  // C: Info Value / part of TÊN VĂN BẢN
                        { wch: 20 },  // D: SỐ VB
                        { wch: 15 },  // E: NGÀY VB
                        { wch: 15 },  // F: FILE VB
                        { wch: 15 },  // G: CQBH
                        { wch: 20 },  // H: GHI CHÚ
                      ];
                      
                      // Cài đặt font cho Info Key & Info Value (các dòng trước Header)
                      for (let r = 0; r < headerRowIndex; r++) {
                        for (let c = 0; c <= 7; c++) {
                           const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
                           if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
                           ws[cellRef].s = {
                             font: { name: 'Times New Roman', sz: 11 },
                             alignment: { vertical: 'center' }
                           };
                           
                           if (r === 0 && c === 0) { // DANH MỤC HỒ SƠ
                               ws[cellRef].s.font.bold = true;
                               ws[cellRef].s.font.sz = 14;
                               ws[cellRef].s.alignment.horizontal = 'center';
                           } else if ((c === 0 || c === 2) && r >= 2) {
                             ws[cellRef].s.alignment.wrapText = true;
                           }
                        }
                      }

                      // Áp dụng format (Border, Center, Bold cho Header, Wrap cho B & C)
                      for (let r = headerRowIndex; r < wsData.length; r++) {
                        for (let c = 0; c <= 7; c++) {
                           const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
                           if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
                           ws[cellRef].s = {
                             border: borderObj,
                             font: { name: 'Times New Roman', sz: 11 },
                             alignment: { vertical: 'center' }
                           };
                           
                           // Ensure text format to prevent Excel from mangling fractions or dates
                           ws[cellRef].z = '@';
                           
                           if (r === headerRowIndex) { // Header
                             ws[cellRef].s.font.bold = true;
                             ws[cellRef].s.alignment.horizontal = 'center';
                           } else { // Data
                             if (c === 0 || c === 3 || c === 4 || c === 5 || c === 6) { // Thêm cột 6 (G)
                                ws[cellRef].s.alignment.horizontal = 'center';
                                if (c === 3) {
                                  ws[cellRef].s.alignment.wrapText = true; // Cho phép xuống dòng ở Số VB
                                }
                             } else if (c === 1 || c === 2) {
                                ws[cellRef].s.alignment.wrapText = true;
                             }
                           }
                        }
                      }
                      
                      // Set file links
                      data.forEach((row, idx) => {
                         let fileLink = row[5];
                         if (fileLink && (fileLink.startsWith('http://') || fileLink.startsWith('https://'))) {
                            const r = headerRowIndex + 1 + idx;
                            const cellRef = XLSX.utils.encode_cell({ r: r, c: 5 });
                            if (ws[cellRef]) {
                               ws[cellRef].l = { Target: fileLink };
                               ws[cellRef].s.font.color = { rgb: '0000FF' };
                               ws[cellRef].s.font.underline = true;
                            }
                         }
                      });

                      XLSX.utils.book_append_sheet(wb, ws, "Danh mục hồ sơ");
                      const projectCode = projectInfo["Mã công trình"] || "Unknown";
                      const filename = `Danh mục hồ sơ - ${projectCode}.xlsx`;
                      XLSX.writeFile(wb, filename);
                    } catch (error) {
                      console.error('Error generating Excel file:', error);
                      alert('Có lỗi xảy ra khi tạo file Excel. Vui lòng thử lại.');
                    }
                  }}
                  className="flex flex-row items-center justify-center gap-1.5 px-3 py-2 bg-white text-emerald-700 border border-emerald-200 rounded-lg hover:bg-emerald-50 font-bold shadow-sm hover:shadow-md active:scale-95 transition-all duration-200 cursor-pointer w-full"
                >
                  <Download className="w-4 h-4" />
                  <span className="uppercase tracking-wider text-[11px] sm:text-xs">Tải xuống</span>
                </button>
              </div>
            </div>

            {/* Project Info */}
            {Object.keys(projectInfo).length > 0 && (
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-col flex-1">
                <h2 className="text-base font-bold text-slate-800 mb-4 border-b border-slate-100 pb-2 uppercase tracking-wide">Thông tin dự án</h2>
                <div className="flex flex-col gap-y-4">
                  {Object.entries(projectInfo).map(([key, value], idx) => {
                    if (key === 'Tên dự án/công trình') {
                      return (
                        <div key={idx} className="flex flex-col gap-1.5">
                          <span className="text-slate-500 font-medium text-xs uppercase tracking-wide">{key}</span>
                          <div className="relative w-full" ref={dropdownRef}>
                            {!showDropdown && value ? (
                              <div 
                                className="w-full px-3 py-2.5 min-h-[42px] border border-transparent hover:border-slate-300 rounded-md bg-slate-50 cursor-pointer text-slate-800 font-semibold text-sm break-words transition-colors flex items-start justify-between gap-2 group"
                                onClick={() => {
                                  setProjectSearchTerm('');
                                  setShowDropdown(true);
                                  setTimeout(() => {
                                    projectNameRef.current?.focus();
                                  }, 10);
                                }}
                              >
                                <span>{value}</span>
                                <button
                                  className="text-slate-400 hover:text-slate-600 p-0.5 rounded shrink-0 opacity-0 group-hover:opacity-100 transition-opacity"
                                  title="Xoá"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setProjectInfo(prev => ({...prev, "Tên dự án/công trình": ""}));
                                    setProjectSearchTerm('');
                                    setShowDropdown(true);
                                    setTimeout(() => {
                                      projectNameRef.current?.focus();
                                    }, 10);
                                  }}
                                >
                                  <X className="w-4 h-4" />
                                </button>
                              </div>
                            ) : (
                              <textarea 
                                ref={projectNameRef}
                                value={projectSearchTerm}
                                onFocus={() => {
                                  setProjectSearchTerm('');
                                  setShowDropdown(true);
                                }}
                                onInput={(e) => {
                                  const target = e.target as HTMLTextAreaElement;
                                  target.style.height = 'auto';
                                  target.style.height = `${target.scrollHeight}px`;
                                }}
                                onChange={(e) => {
                                  setProjectSearchTerm(e.target.value);
                                }}
                                className="w-full pr-10 pl-3 py-2.5 border border-emerald-500 rounded-md focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none text-slate-800 font-semibold text-sm resize-none overflow-hidden bg-white shadow-sm"
                                placeholder="Nhập từ khoá để tìm kiếm..."
                                rows={1}
                              />
                            )}
                            
                            {showDropdown && projectSearchTerm && (
                              <button 
                                onClick={() => {
                                  setProjectSearchTerm('');
                                }}
                                className="absolute right-3 top-2.5 p-1 text-slate-400 hover:text-slate-600"
                                title="Xoá"
                              >
                                <X className="w-4 h-4" />
                              </button>
                            )}
                            
                            {showDropdown && (
                              <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-md shadow-xl max-h-80 overflow-y-auto w-[calc(100vw-32px)] sm:w-full">
                                {availableProjects
                                  .filter(p => !projectSearchTerm || 
                                    String(p["Tên dự án/công trình"]).toLowerCase().includes(projectSearchTerm.toLowerCase()) || 
                                    String(p["Mã công trình"]).toLowerCase().includes(projectSearchTerm.toLowerCase())
                                  )
                                  .map((p, i) => (
                                    <div 
                                      key={i}
                                      className="px-4 py-2 hover:bg-emerald-50 cursor-pointer border-b border-slate-50 last:border-0 leading-relaxed"
                                      onClick={() => {
                                        setProjectInfo(p);
                                        setShowDropdown(false);
                                        setProjectSearchTerm('');
                                      }}
                                    >
                                      <div className="font-semibold text-sm text-slate-800">{p["Tên dự án/công trình"]}</div>
                                      <div className="text-xs text-slate-500 mt-0.5 font-mono">Mã CT: {p["Mã công trình"]}</div>
                                    </div>
                                  ))}
                                {availableProjects.filter(p => !projectSearchTerm || 
                                    String(p["Tên dự án/công trình"]).toLowerCase().includes(projectSearchTerm.toLowerCase()) || 
                                    String(p["Mã công trình"]).toLowerCase().includes(projectSearchTerm.toLowerCase())
                                  ).length === 0 && (
                                  <div className="px-4 py-3 text-sm text-slate-400 italic">Không tìm thấy dự án phù hợp</div>
                                )}
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    }
                    return (
                      <div key={idx} className="flex flex-col gap-1">
                        <span className="text-slate-500 font-medium text-xs uppercase tracking-wide">{key}</span>
                        <span className="text-slate-800 font-semibold text-sm">{value}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          {/* MAIN CONTENT RIGHT (Desktop) */}
          <div className="flex-1 w-full min-w-0 flex flex-col">
            {/* Data Table */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-full">
          <div className="px-6 py-4 border-b border-slate-200 bg-white flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div className="flex justify-between items-center gap-4 w-full lg:w-auto">
              <h2 className="font-semibold text-slate-800 shrink-0 uppercase tracking-wider text-sm">DANH MỤC HỒ SƠ</h2>
              
              <div className="flex items-center gap-2 lg:hidden">
                <span className="text-[10px] font-medium text-slate-500 bg-slate-100 px-2 py-1 rounded-full whitespace-nowrap">
                  {visibleRows.length} dòng
                </span>
                <button
                  onClick={() => {
                    if (!scriptUrl) {
                      alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                      setShowSettings(true);
                      return;
                    }
                    if (!projectInfo["Mã công trình"]) {
                      alert('Vui lòng chọn công trình trước khi thêm nội dung!');
                      return;
                    }
                    setShowAddContentModal(true);
                  }}
                  className="p-1 px-[5px] bg-slate-100 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-md transition-colors"
                  title="Thêm nội dung văn bản"
                >
                  <Plus className="w-[14px] h-[14px] stroke-[3]" />
                </button>
                <button 
                  onClick={() => setShowMobileFilters(!showMobileFilters)}
                  className={`p-2 rounded-lg border transition-all relative ${
                    showMobileFilters || selectedSection !== 'all' 
                      ? 'bg-emerald-50 border-emerald-200 text-emerald-600' 
                      : 'bg-white border-slate-200 text-slate-400 hover:text-slate-600'
                  }`}
                  title="Lọc theo mục"
                >
                  <Filter className="w-4 h-4" />
                  {selectedSection !== 'all' && (
                    <span className="absolute top-1 right-1 w-2 h-2 bg-emerald-500 rounded-full border border-white"></span>
                  )}
                </button>
              </div>
            </div>

            {/* Search and Filters Area */}
            <div className={`${showMobileFilters ? 'flex' : 'hidden'} lg:flex flex-col lg:flex-row lg:items-center gap-3 w-full lg:w-auto`}>
              {/* Section Filter - Toggleable on mobile, always visible on desktop */}
              <div className="relative w-full lg:w-64">
                <select
                  value={selectedSection}
                  onChange={(e) => setSelectedSection(e.target.value)}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none bg-white appearance-none pr-10"
                >
                  <option value="all">Tất cả các mục</option>
                  {sections.map((section, idx) => (
                    <option key={idx} value={section}>{section}</option>
                  ))}
                </select>
                <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
              </div>

              {/* Search - Always visible on desktop, toggleable with filters on mobile if needed, but user wanted it outside on mobile before */}
              {/* Actually, let's keep search always visible as per previous request, but align it nicely */}
              <div className="relative w-full lg:w-64">
                <input
                  type="text"
                  placeholder="Tìm kiếm nội dung..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="px-3 pr-9 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none w-full"
                />
                {searchTerm && (
                  <button 
                    onClick={() => setSearchTerm('')}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-600 transition-colors"
                    title="Xoá tìm kiếm"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>

              <div className="hidden lg:flex items-center gap-2">
                <span className="text-xs font-medium text-slate-500 bg-slate-200 px-2.5 py-1 rounded-full whitespace-nowrap">
                  {visibleRows.length} dòng
                </span>
                <button
                  onClick={() => {
                    if (!scriptUrl) {
                      alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                      setShowSettings(true);
                      return;
                    }
                    if (!projectInfo["Mã công trình"]) {
                      alert('Vui lòng chọn công trình trước khi thêm nội dung!');
                      return;
                    }
                    setShowAddContentModal(true);
                  }}
                  className="p-1 px-1.5 bg-slate-200 text-slate-500 hover:text-emerald-600 hover:bg-emerald-50 rounded-md transition-colors"
                  title="Thêm nội dung văn bản"
                >
                  <Plus className="w-[14px] h-[14px] stroke-[3]" />
                </button>
              </div>
            </div>
          </div>
          
          {loading ? (
            <div className="p-12 flex flex-col items-center justify-center text-slate-500">
              <RefreshCw className="w-8 h-8 animate-spin mb-4 text-emerald-600" />
              <p>Đang tải dữ liệu ...</p>
            </div>
          ) : error ? (
            <div className="p-8 text-center text-red-600">
              <AlertCircle className="w-12 h-12 mx-auto mb-3 opacity-50" />
              <p className="font-medium">Lỗi khi tải dữ liệu</p>
              <p className="text-sm mt-1">{error}</p>
            </div>
          ) : (
            <>
              {/* Desktop View */}
              <div className="hidden md:block overflow-x-auto max-h-[600px]">
                <table className="w-full text-sm text-left whitespace-nowrap">
                  <thead className="text-xs text-slate-600 uppercase bg-white sticky top-0 z-10 shadow-sm text-center">
                    <tr>
                      {/* Custom headers based on the CSV structure we saw */}
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">STT</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">TÊN VĂN BẢN</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">SỐ VB</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">NGÀY VB</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">FILE VB</th>
                      <th className="px-4 py-3 font-semibold border-b border-slate-200">CƠ QUAN BAN HÀNH</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {visibleRows.map((row, rowIndex) => {
                      const typeData = row[2] || '';
                      const isGD = typeData.startsWith('TYPE:GD|');
                      const isDM = typeData.startsWith('TYPE:DM|');
                      const isSectionHeader = isGD || isDM;
                      
                      let toggleId = '';
                      if (isGD) toggleId = `GD|${typeData.split('|')[1]}`;
                      else if (isDM) toggleId = `DM|${typeData.split('|')[1]}|${typeData.split('|')[2]}`;
                      
                      const isCollapsed = isSectionHeader && collapsedSections.has(toggleId);
                      
                      return (
                        <tr 
                          key={rowIndex} 
                          className={`group hover:bg-slate-50 transition-colors ${
                            isGD ? 'bg-emerald-100/50 font-bold text-emerald-900 cursor-pointer' : 
                            isDM ? 'bg-slate-50 font-semibold text-emerald-800 cursor-pointer' : ''
                          }`}
                          onClick={() => isSectionHeader && toggleSection(toggleId)}
                        >
                          <td className="px-4 py-3 border-r border-slate-100 text-center text-slate-500">
                            {isSectionHeader ? (
                              <div className="flex items-center justify-center">
                                <ChevronDown className={`w-4 h-4 transition-transform duration-200 ${isCollapsed ? '-rotate-90' : ''}`} />
                              </div>
                            ) : row[0]}
                          </td>
                          <td className={`px-4 py-3 border-r border-slate-100 whitespace-normal min-w-[300px] text-slate-700 ${isDM ? 'pl-8' : ''}`}>
                            {isDM ? (
                              <div className="flex items-center justify-between gap-2 relative">
                                <span>{row[1]}</span>
                                <div className={`flex items-center justify-center transition-opacity ${activeDropdownMenu === `desktop-${toggleId}` ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'}`} ref={activeDropdownMenu === `desktop-${toggleId}` ? dropdownMenuRef : null}>
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      setActiveDropdownMenu(activeDropdownMenu === `desktop-${toggleId}` ? null : `desktop-${toggleId}`);
                                    }}
                                    className={`p-1.5 rounded-md transition-colors ${activeDropdownMenu === `desktop-${toggleId}` ? 'bg-emerald-100 text-emerald-700' : 'hover:bg-slate-200 text-slate-400 hover:text-slate-600'}`}
                                    title="Tùy chọn"
                                  >
                                    <MoreVertical className="w-4 h-4" />
                                  </button>
                                  {activeDropdownMenu === `desktop-${toggleId}` && (
                                    <div className="absolute right-0 top-full mt-1 w-max bg-white border border-slate-200 rounded-lg shadow-xl z-50 overflow-hidden py-1">
                                      <button 
                                        type="button"
                                        className="w-full px-4 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 text-sm text-slate-700 font-medium whitespace-nowrap text-left"
                                        onClick={(e) => {
                                          e.preventDefault();
                                          e.stopPropagation();
                                          setActiveDropdownMenu(null);
                                          
                                          if (!scriptUrl) {
                                            alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                                            setShowSettings(true);
                                            return;
                                          }
                                          if (!projectInfo["Mã công trình"]) {
                                            alert('Vui lòng chọn công trình trước khi thêm nội dung!');
                                            return;
                                          }

                                          const gdMa = typeData.split('|')[1];
                                          const dmMa = typeData.split('|')[2];
                                          setAddContentForm({
                                            giaiDoan: gdMa,
                                            dauMuc: dmMa,
                                            noiDung: '',
                                            phongBan: 'QLĐT'
                                          });
                                          setShowAddContentModal(true);
                                        }}
                                      >
                                        <FilePlus className="w-4 h-4 text-slate-400" />
                                        Thêm nội dung văn bản
                                      </button>
                                    </div>
                                  )}
                                </div>
                              </div>
                            ) : (
                               row[1]
                            )}
                          </td>
                          <td className="px-4 py-3 border-r border-slate-100 text-center">{!isSectionHeader ? row[3] : ''}</td>
                          <td className="px-4 py-3 border-r border-slate-100">{!isSectionHeader ? row[4] : ''}</td>
                          <td className="px-4 py-3 border-r border-slate-100">
                            {!isSectionHeader && row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')) ? (
                              <a 
                                href={row[5]} 
                                target="_blank" 
                                rel="noopener noreferrer"
                                className="inline-flex items-center gap-1.5 text-indigo-600 hover:text-indigo-800 hover:bg-indigo-50 px-2 py-1 rounded-md transition-all font-medium group"
                              >
                                <Eye className="w-4 h-4 group-hover:scale-110 transition-transform" />
                                <span>Xem file</span>
                              </a>
                            ) : (
                              !isSectionHeader && <span className={row[5]?.trim().toLowerCase() === 'xem file' ? 'text-slate-400 italic' : ''}>
                                {row[5]}
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-3 border-r border-slate-100 text-center">{!isSectionHeader ? row[6] : ''}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Mobile View (Module/Card Layout) */}
              <div className="md:hidden divide-y divide-slate-100">
                {visibleRows.map((row, rowIndex) => {
                  const typeData = row[2] || '';
                  const isGD = typeData.startsWith('TYPE:GD|');
                  const isDM = typeData.startsWith('TYPE:DM|');
                  const isSectionHeader = isGD || isDM;
                  
                  let toggleId = '';
                  if (isGD) toggleId = `GD|${typeData.split('|')[1]}`;
                  else if (isDM) toggleId = `DM|${typeData.split('|')[1]}|${typeData.split('|')[2]}`;
                  
                  const isCollapsed = isSectionHeader && collapsedSections.has(toggleId);
                  
                  if (isSectionHeader) {
                    return (
                      <div 
                        key={rowIndex} 
                        className={`${isGD ? 'bg-emerald-100/50 text-emerald-900 border-b-2 border-emerald-200' : 'bg-slate-50 text-emerald-800 pl-8'} px-4 py-3 font-bold text-xs uppercase tracking-wider flex items-center justify-between cursor-pointer`}
                        onClick={() => toggleSection(toggleId)}
                      >
                        <div className="flex items-center gap-2 max-w-[85%]">
                          <span>{row[1]}</span>
                        </div>
                        <div className="flex items-center gap-3">
                          {isDM && (
                            <div className="relative" ref={activeDropdownMenu === `mobile-${toggleId}` ? dropdownMenuRef : null}>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setActiveDropdownMenu(activeDropdownMenu === `mobile-${toggleId}` ? null : `mobile-${toggleId}`);
                                }}
                                className={`p-1.5 rounded-md transition-colors ${activeDropdownMenu === `mobile-${toggleId}` ? 'bg-emerald-200 text-emerald-800' : 'hover:bg-slate-200 text-emerald-700/60'}`}
                              >
                                <MoreVertical className="w-4 h-4" />
                              </button>
                              {activeDropdownMenu === `mobile-${toggleId}` && (
                                <div className="absolute right-0 top-full mt-1 w-max bg-white border border-slate-200 rounded-lg shadow-xl z-50 overflow-hidden py-1">
                                  <button 
                                    type="button"
                                    className="w-full px-4 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 text-sm text-slate-700 font-medium normal-case whitespace-nowrap text-left"
                                    onClick={(e) => {
                                      e.preventDefault();
                                      e.stopPropagation();
                                      setActiveDropdownMenu(null);

                                      if (!scriptUrl) {
                                        alert('Vui lòng cấu hình Apps Script Web App URL trong phần Cài đặt trước khi thực hiện chức năng này.');
                                        setShowSettings(true);
                                        return;
                                      }
                                      if (!projectInfo["Mã công trình"]) {
                                        alert('Vui lòng chọn công trình trước khi thêm nội dung!');
                                        return;
                                      }

                                      const gdMa = typeData.split('|')[1];
                                      const dmMa = typeData.split('|')[2];
                                      setAddContentForm({
                                        giaiDoan: gdMa,
                                        dauMuc: dmMa,
                                        noiDung: '',
                                        phongBan: 'QLĐT'
                                      });
                                      setShowAddContentModal(true);
                                    }}
                                  >
                                    <FilePlus className="w-4 h-4 text-slate-400" />
                                    Thêm nội dung văn bản
                                  </button>
                                </div>
                              )}
                            </div>
                          )}
                          <ChevronDown className={`w-4 h-4 transition-transform duration-200 shrink-0 ${isCollapsed ? '-rotate-90' : ''}`} />
                        </div>
                      </div>
                    );
                  }

                  const hasNoInfo = !row[3]?.trim() && !row[4]?.trim() && !(row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')));
                  
                  if (hasNoInfo) {
                    return (
                      <div key={rowIndex} className="p-4 bg-white">
                        <div className="flex items-start gap-3">
                          <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded shrink-0 mt-0.5">
                            {row[0]}
                          </span>
                          <div className="flex flex-col gap-1">
                            <h3 className="text-sm font-medium text-slate-700 leading-snug">
                              {row[1]}
                            </h3>
                            <span className="text-[11px] text-slate-400 italic">Chưa có văn bản</span>
                          </div>
                        </div>
                      </div>
                    );
                  }

                  return (
                    <div key={rowIndex} className="p-4 bg-white active:bg-slate-50 transition-colors">
                      <div className="flex items-start gap-3 mb-3">
                        <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded shrink-0 mt-0.5">
                          {row[0]}
                        </span>
                        <h3 className="text-sm font-medium text-slate-700 leading-snug">
                          {row[1]}
                        </h3>
                      </div>
                      
                      <div className="grid grid-cols-2 gap-4 mb-4 ml-8">
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] uppercase text-slate-400 font-bold tracking-widest">Số VB</span>
                          <span className="text-xs font-semibold text-slate-600">{row[3] || '---'}</span>
                        </div>
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] uppercase text-slate-400 font-bold tracking-widest">Ngày VB</span>
                          <span className="text-xs font-semibold text-slate-600">{row[4] || '---'}</span>
                        </div>
                      </div>

                      <div className="ml-8 flex items-end justify-between pt-3 border-t border-slate-50">
                        <div className="flex flex-col gap-0.5">
                          <span className="text-[9px] uppercase text-slate-400 font-bold tracking-widest">Cơ quan ban hành</span>
                          <span className="text-xs text-slate-500 font-medium">{row[6] || '---'}</span>
                        </div>
                        <div>
                          {row[5] && (row[5].startsWith('http://') || row[5].startsWith('https://')) ? (
                            <a 
                              href={row[5]} 
                              target="_blank" 
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 border border-indigo-100 rounded-lg text-[10px] font-bold shadow-sm active:scale-95 transition-all"
                            >
                              <Eye className="w-3.5 h-3.5" />
                              XEM FILE
                            </a>
                          ) : (
                            <span className="text-[10px] text-slate-400 italic font-medium">
                              {row[5] || 'Không có file'}
                            </span>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
                {rows.length === 0 && (
                  <div className="p-12 text-center text-slate-400 italic">
                    Không tìm thấy dữ liệu phù hợp
                  </div>
                )}
              </div>
            </>
          )}
        </div>
          </div>
        </div>
      </main>

      {/* Action Loading Overlay */}
      {actionLoading && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center">
          <div className="bg-white rounded-xl shadow-xl p-8 max-w-sm w-full mx-4 flex flex-col items-center text-center">
            <div className="w-12 h-12 border-4 border-emerald-100 border-t-emerald-600 rounded-full animate-spin mb-4"></div>
            <h3 className="text-lg font-semibold text-slate-800 mb-2">Đang xử lý dữ liệu...</h3>
            <p className="text-slate-500 text-sm">
              Hệ thống đang cập nhật...
            </p>
          </div>
        </div>
      )}
      <footer className="bg-blue-900 text-white/80 py-4 mt-12">
        <div className="max-w-7xl mx-auto px-4 text-center text-xs font-medium tracking-wide">
          <p>© 2026 Công ty Điện lực Vũng Tàu</p>
        </div>
      </footer>
    </div>
  );
}
