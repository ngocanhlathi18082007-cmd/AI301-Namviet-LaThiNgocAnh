(function () {
  "use strict";

  var state = {
    activeTab: "tab-dashboard",
    activeSubTab: "dash-overview",
    sourceName: "Báo cáo doanh thu (50 dòng)",
    headers: [],
    records: [],
    mapping: {},
    xlsxWorkbook: null,
    pipelineResult: null,
    selectedClusterId: null,
    filterStage: "ALL",
    filterTier: "ALL",
    searchQuery: "",
    sortBy: "default",
    sortDir: "desc",
    isFinalized: false,
    diagnostics: {
      issues: [],
      cellIssueMap: {},
      errorCount: 0,
      warningCount: 0,
      fixedCount: 0,
      validCount: 0
    },
    llmConfig: {
      provider: "offline",
      model: "gemini-2.5-flash",
      apiKey: ""
    }
  };

  var CANONICAL_FIELDS = [
    { key: "lead_id", label: "Mã Lead (lead_id)" },
    { key: "created_at", label: "Ngày tạo (created_at)" },
    { key: "source", label: "Nguồn (source)" },
    { key: "company_name", label: "Tên doanh nghiệp *" },
    { key: "tax_id", label: "Mã số thuế (tax_id)" },
    { key: "contact_name", label: "Người liên hệ" },
    { key: "contact_role", label: "Chức vụ" },
    { key: "email", label: "Email" },
    { key: "phone", label: "Số điện thoại" },
    { key: "region", label: "Khu vực / Tỉnh thành" },
    { key: "industry", label: "Ngành nghề" },
    { key: "company_size", label: "Quy mô nhân sự" },
    { key: "need_description", label: "Nội dung nhu cầu *" },
    { key: "budget_vnd", label: "Ngân sách (VND)" },
    { key: "timeline", label: "Thời hạn triển khai" }
  ];

  var DEFAULT_HEADERS = [
    "lead_id", "created_at", "source", "company_name", "tax_id",
    "contact_name", "contact_role", "email", "phone", "region",
    "industry", "company_size", "need_description", "budget_vnd", "timeline"
  ];

  var SAMPLE_DIRTY_LEADS = [
    {
      lead_id: "LD-ERR-01",
      created_at: "2026-11-15 08:30",
      source: "Hội thảo Chuyển đổi số",
      company_name: "Công ty Cổ phần Công nghệ Ánh Dương",
      tax_id: "0108899888",
      contact_name: "Nguyễn Văn Hùng",
      contact_role: "Giám đốc CNTT",
      email: "hung.nguyen@anhduongtech.vn",
      phone: "0912345678",
      region: "Hà Nội",
      industry: "Công nghệ thông tin",
      company_size: 65,
      need_description: "Cần tích hợp ERP Quản trị Sản xuất và kết nối CRM.",
      budget_vnd: 90000000,
      timeline: "Q1/2027"
    },
    {
      lead_id: "LD-ERR-02",
      created_at: "2026-11-15 09:15",
      source: "Website Form",
      company_name: "Tập đoàn Đại Nam Corp (Trùng MST với dòng 1)",
      tax_id: "0108899888",
      contact_name: "Trần Thị Bích",
      contact_role: "Kế toán trưởng",
      email: "bich.tran@dainam.com.vn",
      phone: "0987654321",
      region: "Hà Nội",
      industry: "Bán lẻ",
      company_size: 120,
      need_description: "Phần mềm kế toán đa chi nhánh.",
      budget_vnd: 60000000,
      timeline: "Tháng sau"
    },
    {
      lead_id: "LD-ERR-03",
      created_at: "2026-11-15 10:00",
      source: "Campaign Q4",
      company_name: "Doanh nghiệp Tư nhân Vận tải Sao Mai",
      tax_id: "0314567890",
      contact_name: "Lê Hoàng Phúc",
      contact_role: "Trưởng phòng Vận hành",
      email: "phuc.le@saomai.vn",
      phone: "0903-abc-789",
      region: "TP.HCM",
      industry: "Logistics & Vận tải",
      company_size: 45,
      need_description: "Hệ thống quản lý định vị đội xe và tối ưu chi phí kho.",
      budget_vnd: 45000000,
      timeline: "Gấp trong tuần"
    },
    {
      lead_id: "LD-ERR-04",
      created_at: "2026-11-15 11:20",
      source: "Hội chợ Thương mại",
      company_name: "Tập đoàn Nông nghiệp & Dược liệu Phúc Lộc Thọ",
      tax_id: "0401239876",
      contact_name: "Phạm Quốc Tuấn",
      contact_role: "Phó Tổng Giám đốc",
      email: "contact@phucloctho",
      phone: "0934112233",
      region: "Đà Nẵng",
      industry: "Dược phẩm",
      company_size: 210,
      need_description: "Phần mềm quản lý chất lượng GMP và theo dõi lô hàng xuất khẩu.",
      budget_vnd: 150000000,
      timeline: "Q1/2027"
    },
    {
      lead_id: "LD-ERR-05",
      created_at: "2026-11-15 14:10",
      source: "Zalo OA",
      company_name: "Chi nhánh XNK Thủy sản An Phát Cần Thơ",
      tax_id: "1800554433",
      contact_name: "Huỳnh Minh Trí",
      contact_role: "Giám đốc Điều hành",
      email: "tri.huynh@anphatseafood.com",
      phone: "0945998877",
      region: "Cần Thơ",
      industry: "Chế biến & Thủy sản",
      company_size: 180,
      need_description: "Triển khai hệ thống truy xuất nguồn gốc đơn hàng.",
      budget_vnd: -45000000,
      timeline: "Trong tháng"
    },
    {
      lead_id: "LD-ERR-06",
      created_at: "2026-11-15 15:00",
      source: "Giới thiệu đối tác",
      company_name: "Công ty TNHH Cơ khí & Đóng tàu Việt Đức",
      tax_id: "030999",
      contact_name: "Đỗ Anh Tuấn",
      contact_role: "Chủ tịch HĐQT",
      email: "tuan.vietduc@gmail.com",
      phone: "0908123456",
      region: "Hải Phòng",
      industry: "Cơ khí chế tạo",
      company_size: 95,
      need_description: "Giải pháp ERP tính giá thành đơn hàng gia công kim loại.",
      budget_vnd: 85000000,
      timeline: "Ngay khi có báo giá"
    },
    {
      lead_id: "LD-ERR-07",
      created_at: "2026-11-15 15:45",
      source: "Google Search",
      company_name: "",
      tax_id: "0209876543",
      contact_name: "Vũ Đình Khiêm",
      contact_role: "Phụ trách Mua hàng",
      email: "khiem.vu@supplier.vn",
      phone: "0977889900",
      region: "Bắc Ninh",
      industry: "Điện tử",
      company_size: 35,
      need_description: "Tìm phần mềm CRM theo dõi báo giá khách mua buôn.",
      budget_vnd: 35000000,
      timeline: "Tuần tới"
    },
    {
      lead_id: "LD-ERR-08",
      created_at: "2026-11-15 16:30",
      source: "Hội thảo Hà Nội",
      company_name: "Công ty CP Bánh kẹo & Thực phẩm Hữu Nghị",
      tax_id: "0109988776",
      contact_name: "Bùi Thị Mai",
      contact_role: "Giám đốc Kinh doanh",
      email: "mai.bui@huunghi.vn",
      phone: "0966554433",
      region: "Hà Nội",
      industry: "Thực phẩm F&B",
      company_size: 350,
      need_description: "Hệ thống DMS quản lý tuyến bán hàng cho 120 đại lý.",
      budget_vnd: "Thương lượng sau",
      timeline: "Q2/2027"
    },
    {
      lead_id: "LD-ERR-09",
      created_at: "2026-11-15 17:15",
      source: "Website Direct",
      company_name: "Công ty Thiết bị Y tế Tân Bình",
      tax_id: "0307766554",
      contact_name: "Hoàng Văn Nam",
      contact_role: "Phó GĐ Kinh doanh",
      email: "nam.hoang@tanbinhmed.vn",
      phone: "0918776655",
      region: "Tokyo / Hải ngoại",
      industry: "Thiết bị y tế",
      company_size: 80,
      need_description: "CRM quản lý phân phối vật tư bệnh viện tuyến tỉnh.",
      budget_vnd: 120000000,
      timeline: "Q1/2027"
    },
    {
      lead_id: "LD-ERR-10",
      created_at: "2026-11-15 18:00",
      source: "Hội thảo Doanh nghiệp",
      company_name: "Công ty Cổ phần Xây dựng Hòa Bình Gia Lai",
      tax_id: "5901234567",
      contact_name: "Ngô Quốc Thịnh",
      contact_role: "Giám đốc Dự án",
      email: "thinh.ngo@hoabinhgialai.vn",
      phone: "0905123999",
      region: "Gia Lai",
      industry: "Xây dựng",
      company_size: 140,
      need_description: "Phần mềm quản trị dòng tiền dự án thầu xây dựng.",
      budget_vnd: 95000000,
      timeline: "Trong quý"
    }
  ];

  var STEP_DETAILS = {
    1: {
      badge: "BƯỚC 01 • NẠP FILE",
      en: "Intake & Multi-Format Parsing",
      vn: "Nhận File Excel / Google Sheets",
      desc: "Tự động nhận diện mọi file CSV (UTF-8, ANSI), XLSX nhiều sheet hoặc link Google Sheets. Tự động thích ứng cấu trúc cột bất kỳ.",
      status: "Sẵn sàng tiếp nhận dữ liệu",
      sla: "< 0.1s",
      guardrail: "Tương thích 100% file",
      output: "Raw Tabular Grid",
      contract_in: "Bất kỳ file .csv, .xlsx hoặc Google Sheets link",
      contract_out: "Bảng dữ liệu trích xuất dạng chuẩn headers & records"
    },
    2: {
      badge: "BƯỚC 02 • ĐỐI SOÁT CỘT",
      en: "Schema Validation & Mapping",
      vn: "Kiểm tra Định dạng (Đầu vào hợp lệ?)",
      desc: "Tự động ánh xạ 15 trường thông tin chuẩn (MST, Email, SĐT, Nhu cầu, Ngân sách). Phát hiện lỗi cấu trúc và bù đắp trường thiếu.",
      status: "Đảm bảo không mất dữ liệu",
      sla: "< 0.05s",
      guardrail: "0% Bị từ chối file",
      output: "Mapped Schema 15 Trường",
      contract_in: "Bảng dữ liệu raw với tên cột bất kỳ",
      contract_out: "Ánh xạ chuẩn 15 trường thông tin doanh nghiệp"
    },
    3: {
      badge: "BƯỚC 03 • ĐÃ SẠCH?",
      en: "Normalize & Stabilize",
      vn: "Làm sạch Dữ liệu Đa tầng (Đã sạch & ổn?)",
      desc: "Chuẩn hóa SĐT (+84➔0), ngày YYYY-MM-DD, phân bổ khu vực 3 miền Bắc - Trung - Nam, làm sạch khoảng trắng và tên viết tắt.",
      status: "Sẵn sàng trước khi Process",
      sla: "< 0.2s",
      guardrail: "Chuẩn hóa SĐT & Email",
      output: "Clean Normalized Leads",
      contract_in: "Dữ liệu đã ánh xạ chưa chuẩn hóa format",
      contract_out: "Hồ sơ Leads sạch sẽ 100%, chuẩn format quốc gia"
    },
    4: {
      badge: "BƯỚC 04 • PROCESS",
      en: "AI Agent 5 Skills + Gemini",
      vn: "Process (AI Agent 5 Kỹ năng Điều phối)",
      desc: "Lọc trùng viết tắt (Jaro-Winkler + MST), chặn DNC/Opt-out, chấm điểm ICP 0-100đ, chia việc cân bằng tải cho 8 NVKD.",
      status: "Chống bẫy ngân sách ảo",
      sla: "< 0.5s",
      guardrail: "0% Trùng lặp & Cân bằng 8 Sales",
      output: "Clustered & Scored Deals",
      contract_in: "Danh sách hồ sơ khách hàng đã làm sạch",
      contract_out: "Cụm khách hàng chấm điểm ICP & gán đích danh NVKD"
    },
    5: {
      badge: "BƯỚC 05 • VERIFY & HITL",
      en: "Verification & Dual HITL Gate",
      vn: "Kiểm chứng 5 Tiêu chuẩn & Chốt chặn HITL",
      desc: "Kiểm chứng 5 tiêu chuẩn an toàn; kích hoạt [HITL 1] Trưởng phòng KD duyệt ngoại lệ; kích hoạt [HITL 2] Sales duyệt gửi thông điệp.",
      status: "Kiểm soát rủi ro 100%",
      sla: "< 0.1s",
      guardrail: "2 Lớp Phê duyệt Con người",
      output: "Audit Log & Approved Queue",
      contract_in: "Kết quả phân bổ và đề xuất phản hồi của AI",
      contract_out: "Biên bản kiểm thử tự động + Hàng đợi duyệt Trưởng phòng"
    },
    6: {
      badge: "BƯỚC 06 • BÁO CÁO BI",
      en: "Executive BI Dashboard & Export",
      vn: "Dashboard, Chốt Kết quả & Xuất HTML",
      desc: "Duyệt ca Trưởng phòng, trực quan hóa 5 cụm biểu đồ SVG vector, chốt kết quả và tải báo cáo HTML độc lập kèm Audit Log.",
      status: "Chốt kết quả & xuất file",
      sla: "< 0.3s",
      guardrail: "Xuất HTML chạy Offline 100%",
      output: "5 SVG Charts & Single HTML",
      contract_in: "Toàn bộ cụm deals đã được xác minh & phê duyệt",
      contract_out: "Dashboard trực quan + Báo cáo HTML độc lập"
    }
  };

  function $(id) {
    return document.getElementById(id);
  }

  function escHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function formatCompactVnd(amount) {
    if (!amount || amount <= 0) return "Chưa rõ";
    if (amount >= 1000000000) {
      return (amount / 1000000000).toFixed(2).replace(/\.00$/, "") + " Tỷ đ";
    }
    if (amount >= 1000000) {
      return Math.round(amount / 1000000) + " Tr đ";
    }
    return Number(amount).toLocaleString("en-US") + " đ";
  }

  function showToast(msg, type) {
    if (typeof document === "undefined") return;
    var container = $("globalToastContainer");
    if (!container) {
      container = document.createElement("div");
      container.id = "globalToastContainer";
      container.style.cssText = "position:fixed;bottom:24px;right:24px;z-index:99999;display:flex;flex-direction:column;gap:10px;pointer-events:none;";
      document.body.appendChild(container);
    }
    var toast = document.createElement("div");
    var bg = type === "error" ? "#fef2f2" : "#f0fdf4";
    var border = type === "error" ? "#fecaca" : "#bbf7d0";
    var color = type === "error" ? "#991b1b" : "#166534";
    var ico = type === "error" ? "⚠️" : "✓";
    toast.style.cssText = "pointer-events:auto;min-width:280px;max-width:440px;padding:12px 16px;border-radius:12px;background:" + bg + ";border:1.5px solid " + border + ";color:" + color + ";box-shadow:0 10px 25px rgba(0,0,0,0.1);font-size:13px;font-weight:700;display:flex;align-items:center;gap:10px;";
    toast.innerHTML = "<span>" + ico + "</span><span style='flex:1;line-height:1.4;'>" + escHtml(msg) + "</span>";
    container.appendChild(toast);
    setTimeout(function () {
      toast.style.opacity = "0";
      toast.style.transform = "translateY(10px)";
      toast.style.transition = "all 0.3s ease";
      setTimeout(function () { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 300);
    }, 4000);
  }

  function showStatusBanner(msg, type) {
    var el = $("intakeStatusBanner");
    if (el) {
      el.textContent = msg;
      el.className = "status-banner-box " + (type || "success");
      el.classList.remove("hidden");
    }
    showToast(msg, type);
  }

  function hideStatusBanner() {
    var el = $("intakeStatusBanner");
    if (el) el.classList.add("hidden");
  }

  // =========================================================================
  // 1. Navigation Across the 4 Main Tabs & Subtabs
  // =========================================================================
  function switchPageTab(tabId) {
    state.activeTab = tabId;
    document.querySelectorAll(".nav-tab-pill").forEach(function (pill) {
      pill.classList.toggle("active", pill.dataset.tab === tabId);
    });
    document.querySelectorAll(".app-page-view").forEach(function (view) {
      view.classList.toggle("active", view.id === tabId);
    });
    window.scrollTo({ top: 0, behavior: "smooth" });

    if (tabId === "tab-dashboard") {
      renderDashboard();
    }
  }

  function switchDashSubTab(subTabId) {
    state.activeSubTab = subTabId;
    document.querySelectorAll(".dash-sub-pill").forEach(function (pill) {
      pill.classList.toggle("active", pill.dataset.subtab === subTabId);
    });
    document.querySelectorAll(".dash-subpanel").forEach(function (panel) {
      panel.classList.toggle("active", panel.id === "subtab-" + subTabId);
    });
  }

  function updateStepPreview(stepNum) {
    var data = STEP_DETAILS[stepNum];
    if (!data) return;

    if ($("wfBadgeStep")) $("wfBadgeStep").textContent = data.badge;
    if ($("wfEnTitle")) $("wfEnTitle").textContent = data.en;
    if ($("wfVnTitle")) $("wfVnTitle").textContent = data.vn;
    if ($("wfStepDesc")) $("wfStepDesc").textContent = data.desc;
    if ($("wfFooterStatus")) $("wfFooterStatus").textContent = data.status;
    if ($("wfSpecSla")) $("wfSpecSla").textContent = data.sla || "< 0.2s";
    if ($("wfSpecGuard")) $("wfSpecGuard").textContent = data.guardrail || "100% An toàn";
    if ($("wfSpecOutput")) $("wfSpecOutput").textContent = data.output || "Clean JSON";
    if ($("wfContractIn")) $("wfContractIn").textContent = data.contract_in || "Dữ liệu đầu vào";
    if ($("wfContractOut")) $("wfContractOut").textContent = data.contract_out || "Dữ liệu đầu ra";

    document.querySelectorAll(".wf-step-row-item").forEach(function (row) {
      row.classList.toggle("active", parseInt(row.dataset.step, 10) === stepNum);
    });
    document.querySelectorAll(".wf-ribbon-node").forEach(function (node) {
      node.classList.toggle("active", parseInt(node.dataset.step, 10) === stepNum);
    });
  }

  // =========================================================================
  // 2. Data Ingestion & Quality Diagnostics (CSV, XLSX, Errors, Schema Mapping)
  // =========================================================================
  function inspectRecordsForDiagnostics(headers, records, mapping) {
    mapping = mapping || {};
    var issues = [];
    var cellIssueMap = {};
    var seenTaxIds = {};
    var errorCount = 0;
    var warningCount = 0;
    var fixedCount = 0;

    var colOf = {};
    Object.keys(mapping).forEach(function (k) {
      if (mapping[k]) colOf[k] = mapping[k];
    });

    function addIssue(rowIdx, colHeader, fieldKey, rawVal, severity, reason, fix) {
      var rowNumber = rowIdx + 1;
      var issue = {
        rowNumber: rowNumber,
        colHeader: colHeader,
        fieldKey: fieldKey,
        rawValue: String(rawVal == null ? "" : rawVal).trim(),
        severity: severity,
        reason: reason,
        fix: fix
      };
      issues.push(issue);
      cellIssueMap[rowIdx + "::" + colHeader] = issue;
      if (severity === "error") {
        errorCount++;
      } else {
        warningCount++;
      }
      fixedCount++;
    }

    records.forEach(function (rec, rowIdx) {
      // 1. Company Name Check
      var compCol = colOf["company_name"] || "company_name";
      var compVal = rec[compCol];
      if (compVal == null || String(compVal).trim() === "") {
        addIssue(
          rowIdx,
          compCol,
          "company_name",
          "(Để trống)",
          "error",
          "Tên Doanh nghiệp bị để trống, không thể định danh pháp nhân",
          "Tự động gán nhãn [Doanh nghiệp Khuyết danh - Dòng " + (rowIdx + 1) + "] để không mất cơ hội"
        );
      }

      // 2. Tax ID / MST Check
      var taxCol = colOf["tax_id"] || "tax_id";
      var taxVal = rec[taxCol];
      if (taxVal != null && String(taxVal).trim() !== "") {
        var cleanTax = String(taxVal).trim().replace(/[^0-9]/g, "");
        if (cleanTax.length > 0) {
          if (seenTaxIds[cleanTax] != null) {
            var prevRow = seenTaxIds[cleanTax];
            addIssue(
              rowIdx,
              taxCol,
              "tax_id",
              taxVal,
              "error",
              "Trùng lặp Mã số thuế (" + cleanTax + ") với hồ sơ tại Dòng " + prevRow,
              "Tự động nhóm vào cụm trùng lặp (Duplicate Cluster) và gán cờ HITL 1 phê duyệt"
            );
          } else {
            seenTaxIds[cleanTax] = rowIdx + 1;
          }

          if (cleanTax.length !== 10 && cleanTax.length !== 13) {
            addIssue(
              rowIdx,
              taxCol,
              "tax_id",
              taxVal,
              "warning",
              "Mã số thuế chỉ có " + cleanTax.length + " chữ số (chuẩn Bộ Tài Chính VN là 10 hoặc 13 số)",
              "Giữ nguyên mã thô và gắn nhãn yêu cầu NVKD tra cứu cổng Tổng cục Thuế"
            );
          }
        }
      }

      // 3. Phone Check
      var phoneCol = colOf["phone"] || "phone";
      var phoneVal = rec[phoneCol];
      if (phoneVal != null && String(phoneVal).trim() !== "") {
        var phoneStr = String(phoneVal).trim();
        if (/[a-zA-Z]/.test(phoneStr)) {
          addIssue(
            rowIdx,
            phoneCol,
            "phone",
            phoneVal,
            "error",
            "Số điện thoại chứa ký tự chữ cái không hợp lệ",
            "Tự động loại bỏ ký tự lạ, trích xuất chuỗi số hợp lệ"
          );
        } else {
          var digits = phoneStr.replace(/\D/g, "");
          if (digits.length > 0 && (digits.length < 9 || digits.length > 12)) {
            addIssue(
              rowIdx,
              phoneCol,
              "phone",
              phoneVal,
              "warning",
              "Độ dài số điện thoại bất thường (" + digits.length + " chữ số)",
              "Ghi nhận cờ nghi vấn và đưa vào ghi chú giao việc cho NVKD"
            );
          }
        }
      }

      // 4. Email Check
      var emailCol = colOf["email"] || "email";
      var emailVal = rec[emailCol];
      if (emailVal != null && String(emailVal).trim() !== "") {
        var emailStr = String(emailVal).trim();
        var isValidEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailStr);
        if (!isValidEmail) {
          addIssue(
            rowIdx,
            emailCol,
            "email",
            emailVal,
            "error",
            "Định dạng email không đúng chuẩn RFC (thiếu @ hoặc thiếu tên miền .com / .vn)",
            "Chuyển cờ xác thực [HITL 2] cho NVKD liên hệ qua Zalo/Điện thoại để cập nhật email"
          );
        }
      }

      // 5. Budget / Revenue Check
      var budgetCol = colOf["budget_vnd"] || "budget_vnd";
      var budgetVal = rec[budgetCol];
      if (budgetVal != null && String(budgetVal).trim() !== "") {
        var budgetStr = String(budgetVal).trim();
        var num = Number(String(budgetStr).replace(/,/g, "").replace(/\./g, ""));
        if (!isNaN(num) && num < 0) {
          addIssue(
            rowIdx,
            budgetCol,
            "budget_vnd",
            budgetVal,
            "error",
            "Giá trị ngân sách/doanh số là số âm không hợp lệ (" + budgetStr + ")",
            "Tự động đổi dấu sang giá trị tuyệt đối (" + Math.abs(num).toLocaleString("vi-VN") + "đ) và thông báo Trưởng phòng"
          );
        } else if (isNaN(num) && /[a-zA-Z\u00C0-\u024F\u1E00-\u1EFF]/.test(budgetStr)) {
          addIssue(
            rowIdx,
            budgetCol,
            "budget_vnd",
            budgetVal,
            "warning",
            "Cột số tiền chứa chuỗi văn bản: \"" + budgetStr + "\"",
            "Tạm thời quy đổi 0đ và xếp vào trạng thái [Cần đàm phán ngân sách]"
          );
        }
      }

      // 6. Region Check
      var regionCol = colOf["region"] || "region";
      var regionVal = rec[regionCol];
      if (regionVal != null && String(regionVal).trim() !== "") {
        var regStr = String(regionVal).trim().toLowerCase();
        var hasVnRegion = regStr.indexOf("hà nội") >= 0 || regStr.indexOf("bắc") >= 0 ||
                          regStr.indexOf("hcm") >= 0 || regStr.indexOf("hồ chí minh") >= 0 ||
                          regStr.indexOf("nam") >= 0 || regStr.indexOf("đà nẵng") >= 0 ||
                          regStr.indexOf("trung") >= 0 || regStr.indexOf("cần thơ") >= 0 ||
                          regStr.indexOf("hải phòng") >= 0 || regStr.indexOf("gia lai") >= 0 ||
                          regStr.indexOf("bắc ninh") >= 0;
        if (!hasVnRegion && (regStr.indexOf("tokyo") >= 0 || regStr.indexOf("hải ngoại") >= 0 || regStr.indexOf("nước ngoài") >= 0)) {
          addIssue(
            rowIdx,
            regionCol,
            "region",
            regionVal,
            "warning",
            "Khu vực nằm ngoài địa bàn nội địa tiêu chuẩn (Bắc - Trung - Nam)",
            "Chuyển giao về Đội Khách hàng Doanh nghiệp Quốc tế / Trưởng phòng [HITL 1] phân bổ"
          );
        }
      }
    });

    var totalCells = records.length * (headers.length || 1);
    var validCount = Math.max(0, totalCells - issues.length);

    return {
      issues: issues,
      cellIssueMap: cellIssueMap,
      errorCount: errorCount,
      warningCount: warningCount,
      fixedCount: fixedCount,
      validCount: validCount
    };
  }

  function renderDataQualityDiagnostics() {
    var consoleEl = $("fileQualityDiagnosticConsole");
    if (!consoleEl) return;

    if (state.records.length === 0) {
      consoleEl.classList.add("hidden");
      return;
    }

    consoleEl.classList.remove("hidden");
    var diag = state.diagnostics || { issues: [], errorCount: 0, warningCount: 0, fixedCount: 0, validCount: 0 };

    if ($("diagCountErrors")) $("diagCountErrors").textContent = String(diag.errorCount);
    if ($("diagCountWarnings")) $("diagCountWarnings").textContent = String(diag.warningCount);
    if ($("diagCountFixed")) $("diagCountFixed").textContent = String(diag.fixedCount);
    if ($("diagCountValid")) $("diagCountValid").textContent = String(diag.validCount);

    var tbody = $("diagnosticTableBody");
    var wrapper = $("diagnosticTableWrapper");

    if (diag.issues.length === 0) {
      consoleEl.classList.add("diagnostic-perfect");
      if ($("diagTitleIcon")) $("diagTitleIcon").textContent = "✅";
      if ($("diagTitleText")) $("diagTitleText").textContent = "Tệp Tin Hoàn Toàn Hợp Lệ (0 Điểm Lỗi)";
      if ($("diagSubtitleText")) {
        $("diagSubtitleText").textContent =
          "Toàn bộ " + state.records.length + " dòng dữ liệu và " + state.headers.length +
          " cột đều đạt chuẩn kỹ thuật. Dữ liệu đã sẵn sàng để AI Agent tự động điều phối và tạo Dashboard!";
      }
      if (wrapper) wrapper.classList.add("hidden");
      if (tbody) tbody.innerHTML = "";
    } else {
      consoleEl.classList.remove("diagnostic-perfect");
      if ($("diagTitleIcon")) $("diagTitleIcon").textContent = "🔍";
      if ($("diagTitleText")) {
        $("diagTitleText").textContent =
          "Phát Hiện " + diag.issues.length + " Vị Trí Cần Xử Lý Trong Tệp Tin (" +
          diag.errorCount + " Lỗi định dạng • " + diag.warningCount + " Cảnh báo)";
      }
      if ($("diagSubtitleText")) {
        $("diagSubtitleText").textContent =
          "Hệ thống đã định vị chính xác số dòng và tên cột bị lỗi bên dưới. AI Agent đã áp dụng các giải pháp tự động xử lý để pipeline vẫn chạy thông suốt:";
      }
      if (wrapper) wrapper.classList.remove("hidden");
      if (tbody) {
        tbody.innerHTML = diag.issues.map(function (item) {
          var tagBadge = item.severity === "error"
            ? '<span class="diag-tag diag-tag-danger">❌ Lỗi định dạng</span>'
            : '<span class="diag-tag diag-tag-warning">⚠️ Cảnh báo cấu trúc</span>';
          return (
            "<tr>" +
            '<td><strong style="color:#be123c;">Dòng ' + item.rowNumber + "</strong></td>" +
            '<td><code style="font-size:11.5px;color:#0f172a;font-weight:700;">' + escHtml(item.colHeader) + "</code></td>" +
            '<td><span class="raw-value-badge has-error">' + escHtml(item.rawValue || "(trống)") + "</span></td>" +
            "<td>" + tagBadge + "</td>" +
            '<td style="color:#475569;font-weight:600;">' + escHtml(item.reason) + "</td>" +
            '<td><span class="diag-tag diag-tag-fixed">🤖 ' + escHtml(item.fix) + "</span></td>" +
            "</tr>"
          );
        }).join("");
      }
    }
  }

  function loadTableIntoIntake(headers, records, sourceLabel) {
    state.headers = headers || [];
    state.records = records || [];
    state.sourceName = sourceLabel;
    state.mapping = SpreadsheetIO.detectColumnMapping(state.headers, state.records);

    if ($("intakeRowCount")) $("intakeRowCount").textContent = String(state.records.length);

    var mappedCount = Object.keys(state.mapping).filter(function (k) { return !!state.mapping[k]; }).length;
    if ($("intakeColMappedCount")) $("intakeColMappedCount").textContent = mappedCount + "/15";

    state.diagnostics = inspectRecordsForDiagnostics(state.headers, state.records, state.mapping);
    renderDataQualityDiagnostics();

    renderMappingGrid();
    renderPreviewTable();
  }

  function renderMappingGrid() {
    var grid = $("mappingGrid");
    if (!grid) return;

    grid.innerHTML = CANONICAL_FIELDS.map(function (f) {
      var current = state.mapping[f.key] || "";
      var opts = '<option value="">-- Bỏ trống --</option>' +
        state.headers.map(function (h) {
          return '<option value="' + escHtml(h) + '"' + (h === current ? " selected" : "") + ">" + escHtml(h) + "</option>";
        }).join("");
      return (
        '<div style="display:flex;flex-direction:column;gap:3px;font-size:11.5px;">' +
        '<label style="font-weight:750;color:#57534e;">' + escHtml(f.label) + "</label>" +
        '<select data-field="' + f.key + '" style="border:1px solid #e7e2da;padding:4px 8px;border-radius:6px;background:#fff;font-size:11.5px;">' + opts + "</select>" +
        "</div>"
      );
    }).join("");

    grid.querySelectorAll("select").forEach(function (sel) {
      sel.addEventListener("change", function () {
        state.mapping[sel.dataset.field] = sel.value;
        var mappedCount = Object.keys(state.mapping).filter(function (k) { return !!state.mapping[k]; }).length;
        if ($("intakeColMappedCount")) $("intakeColMappedCount").textContent = mappedCount + "/15";
        state.diagnostics = inspectRecordsForDiagnostics(state.headers, state.records, state.mapping);
        renderDataQualityDiagnostics();
        renderPreviewTable();
      });
    });
  }

  function renderPreviewTable() {
    var table = $("previewTable");
    if (!table) return;
    var thead = table.querySelector("thead");
    var tbody = table.querySelector("tbody");
    if (!thead || !tbody) return;

    if (state.headers.length === 0 || state.records.length === 0) {
      thead.innerHTML = "";
      tbody.innerHTML = '<tr><td style="text-align:center;padding:16px;color:#a8a29e;">Chưa có dữ liệu. Vui lòng tải file hoặc chọn mẫu ở trên.</td></tr>';
      return;
    }

    thead.innerHTML = "<tr><th style='width:50px;'>STT</th>" + state.headers.map(function (h) {
      return "<th>" + escHtml(h) + "</th>";
    }).join("") + "</tr>";

    var slice = state.records.slice(0, 10);
    var cellMap = (state.diagnostics && state.diagnostics.cellIssueMap) ? state.diagnostics.cellIssueMap : {};

    tbody.innerHTML = slice.map(function (rec, rowIdx) {
      var cellsHtml = state.headers.map(function (h) {
        var cellKey = rowIdx + "::" + h;
        var issue = cellMap[cellKey];
        var cellClass = "";
        var icon = "";
        var titleAttr = "";
        if (issue) {
          if (issue.severity === "error") {
            cellClass = ' class="cell-has-error"';
            icon = '<span class="cell-error-icon">❌</span>';
          } else {
            cellClass = ' class="cell-has-warning"';
            icon = '<span class="cell-error-icon">⚠️</span>';
          }
          titleAttr = ' title="[Dòng ' + issue.rowNumber + ' - Cột ' + escHtml(h) + ']: ' + escHtml(issue.reason) + ' ➔ AI: ' + escHtml(issue.fix) + '"';
        }
        return "<td" + cellClass + titleAttr + ">" + icon + escHtml(rec[h] || "") + "</td>";
      }).join("");

      return "<tr><td style='font-weight:700;color:#94a3b8;font-size:11px;'>" + (rowIdx + 1) + "</td>" + cellsHtml + "</tr>";
    }).join("");
  }

  function recordsFromRawObjects(items) {
    return items.map(function (row) {
      var out = {};
      DEFAULT_HEADERS.forEach(function (h) {
        out[h] = row[h] == null ? "" : String(row[h]);
      });
      return out;
    });
  }

  function loadDefaultDemoData() {
    var records = recordsFromRawObjects(NAMVIET_DATA.sampleLeads);
    if ($("sheetSelectorBox")) $("sheetSelectorBox").classList.add("hidden");
    loadTableIntoIntake(DEFAULT_HEADERS.slice(), records, "18 leads mẫu");
  }

  function loadSampleDirtyLeads() {
    var records = recordsFromRawObjects(SAMPLE_DIRTY_LEADS);
    if ($("sheetSelectorBox")) $("sheetSelectorBox").classList.add("hidden");
    loadTableIntoIntake(DEFAULT_HEADERS.slice(), records, "Mẫu có lỗi định dạng & trùng lặp (10 leads)");
    document.querySelectorAll(".intake-presets-row .btn-preset-pill").forEach(function (b) {
      b.classList.remove("active");
    });
    if ($("btnPresetError")) $("btnPresetError").classList.add("active");
    showToast("Đã nạp 10 dòng mẫu có lỗi để kiểm tra hệ thống chẩn đoán!", "warning");
  }

  function loadSampleSalesReport50() {
    var raw = NAMVIET_DATA.sampleSalesReport50 || [];
    if (raw.length === 0) return;
    var headers = Object.keys(raw[0]);
    var records = raw.slice();
    if ($("sheetSelectorBox")) $("sheetSelectorBox").classList.add("hidden");
    loadTableIntoIntake(headers, records, "Báo cáo Doanh thu (50 dòng)");
    document.querySelectorAll(".intake-presets-row .btn-preset-pill").forEach(function (b) {
      b.classList.remove("active");
    });
    if ($("btnPreset50Sales")) $("btnPreset50Sales").classList.add("active");
    runAgentNow({ openDashboard: true });
    showToast("Đã nạp và phân tích thành công 50 dòng từ file báo cáo!", "success");
  }

  function loadSimulatedLargeBatch() {
    var base = NAMVIET_DATA.sampleLeads;
    var generated = [];
    for (var k = 0; k < 120; k++) {
      var src = base[k % base.length];
      var idx = k + 1;
      generated.push({
        lead_id: "LD-" + String(idx).padStart(3, "0"),
        created_at: "2026-11-12 10:00",
        source: src.source,
        company_name: src.company_name + (k >= 18 ? " (Chi nhánh " + (k % 5 + 1) + ")" : ""),
        tax_id: src.tax_id,
        contact_name: "Đầu mối #" + idx,
        contact_role: src.contact_role,
        email: "contact" + idx + "@" + (src.email.split("@")[1] || "gmail.com"),
        phone: src.phone,
        region: src.region,
        industry: src.industry,
        company_size: src.company_size || 50,
        need_description: src.need_description,
        budget_vnd: src.budget_vnd,
        timeline: src.timeline
      });
    }
    var records = recordsFromRawObjects(generated);
    loadTableIntoIntake(DEFAULT_HEADERS.slice(), records, "120 Leads giả lập");
    runAgentNow({ openDashboard: true });
    showToast("Đã sinh và phân tích 120 leads giả lập!", "success");
  }

  function handleUploadedFile(file, autoRunAfterLoad) {
    if (!file) return;
    hideStatusBanner();
    var lower = file.name.toLowerCase();
    var autoRun = autoRunAfterLoad !== undefined ? autoRunAfterLoad : true;

    file.arrayBuffer().then(function (buf) {
      // 1. Try XLSX if named .xlsx or binary
      if (lower.endsWith(".xlsx")) {
        try {
          var wb = SpreadsheetIO.parseXlsxArrayBuffer(buf);
          if (wb && wb.sheetNames && wb.sheetNames.length > 0) {
            state.xlsxWorkbook = wb;
            var select = $("xlsxSheetSelect");
            if (select) {
              select.innerHTML = wb.sheetNames.map(function (name) {
                return '<option value="' + escHtml(name) + '">' + escHtml(name) + " (" + wb.sheets[name].records.length + " dòng)</option>";
              }).join("");
            }
            if ($("sheetSelectorBox")) {
              $("sheetSelectorBox").classList.toggle("hidden", wb.sheetNames.length <= 1);
            }
            // Pick sheet with largest number of records
            var bestSheet = wb.sheetNames.reduce(function (best, cur) {
              return (wb.sheets[cur].records.length > (wb.sheets[best] ? wb.sheets[best].records.length : 0)) ? cur : best;
            }, wb.sheetNames[0]);

            var tbl = wb.sheets[bestSheet];
            loadTableIntoIntake(tbl.headers, tbl.records, "Excel XLSX: " + file.name + " [" + bestSheet + "]");
            showStatusBanner("Đã nạp thành công Excel \"" + file.name + "\" (" + tbl.records.length + " dòng). Tự động phân tích...", "success");
            showToast("Đã nạp file Excel (" + tbl.records.length + " dòng)!", "success");
            if (autoRun) {
              runAgentNow({ openDashboard: true });
            }
            return;
          }
        } catch (_xlsxErr) {
          // XLSX failed, fallback to CSV decode below
        }
      }

      // 2. Decode as Delimited Text (CSV / TSV / Semicolon / Pipe / Plain Text)
      var text = SpreadsheetIO.decodeTextFromBuffer(buf);
      var tbl = SpreadsheetIO.parseCSV(text);

      // If CSV parsed 0 records but might be an un-suffixed XLSX, attempt XLSX
      if ((!tbl || !tbl.records || tbl.records.length === 0) && !lower.endsWith(".xlsx")) {
        try {
          var wbTry = SpreadsheetIO.parseXlsxArrayBuffer(buf);
          if (wbTry && wbTry.sheetNames && wbTry.sheetNames.length > 0) {
            var firstSheet = wbTry.sheetNames[0];
            var tblTry = wbTry.sheets[firstSheet];
            loadTableIntoIntake(tblTry.headers, tblTry.records, "XLSX: " + file.name);
            showStatusBanner("Đã nạp thành công \"" + file.name + "\" (" + tblTry.records.length + " dòng).", "success");
            if (autoRun) runAgentNow({ openDashboard: true });
            return;
          }
        } catch (_e2) {}
      }

      if (!tbl || !tbl.records || tbl.records.length === 0) {
        showStatusBanner("Không tìm thấy dòng dữ liệu trong file \"" + file.name + "\". Hãy kiểm tra nội dung file.", "error");
        return;
      }

      if ($("sheetSelectorBox")) $("sheetSelectorBox").classList.add("hidden");
      loadTableIntoIntake(tbl.headers, tbl.records, "CSV/Text: " + file.name);
      showStatusBanner("Đã nạp thành công \"" + file.name + "\" (" + tbl.records.length + " dòng). Tự động tạo Dashboard...", "success");
      showToast("Đã nạp " + tbl.records.length + " dòng từ file " + file.name + "!", "success");
      if (autoRun) {
        runAgentNow({ openDashboard: true });
      }
    }).catch(function (err) {
      showStatusBanner("Lỗi đọc file: " + err.message, "error");
    });
  }

  // =========================================================================
  // 3. AI Agent Execution Pipeline
  // =========================================================================
  function runAgentNow(options) {
    var opts = options || {};
    if (!state.records || state.records.length === 0) {
      loadSampleSalesReport50();
      return;
    }
    var leads = SpreadsheetIO.mapRecordsToLeads(state.records, state.mapping);
    if (leads.length === 0) {
      showStatusBanner("Không tìm thấy dòng dữ liệu nào để phân tích.", "error");
      return;
    }

    var result = NamVietAgent.runAgentPipeline(leads, NAMVIET_DATA);
    state.pipelineResult = result;
    state.selectedClusterId = result.clusters[0] ? result.clusters[0].cluster_id : null;

    if ($("introLeadsTotal")) $("introLeadsTotal").textContent = String(result.clusters.length);

    if (opts.openDashboard) {
      switchPageTab("tab-dashboard");
      switchDashSubTab("dash-overview");
    } else {
      renderDashboard();
    }
  }

  // =========================================================================
  // 4. Render Executive Dashboard & BI (Professional Detailed Charts)
  // =========================================================================
  function renderDashboardEmptyState() {
    if ($("kpiTotalLeads")) $("kpiTotalLeads").textContent = "--";
    if ($("kpiIcpRate")) $("kpiIcpRate").textContent = "--%";
    if ($("kpiIcpCount")) $("kpiIcpCount").textContent = "Chờ nạp file";
    if ($("kpiTotalBudget")) $("kpiTotalBudget").textContent = "-- đ";
    if ($("kpiDispatchedSales")) $("kpiDispatchedSales").textContent = "0/8 Sales";
    if ($("kpiPendingHitl1")) $("kpiPendingHitl1").textContent = "Chờ nạp file";
    if ($("dashReviewStatusTag")) {
      $("dashReviewStatusTag").textContent = "Chưa có dữ liệu";
      $("dashReviewStatusTag").style.background = "#fff5f8";
      $("dashReviewStatusTag").style.color = "#be185d";
      $("dashReviewStatusTag").style.borderColor = "#fbcfe8";
    }

    var emptyHtml = 
      '<div class="chart-empty-state">' +
      '<div class="empty-icon">📊</div>' +
      '<div class="empty-text">Chưa có dữ liệu</div>' +
      '<div class="empty-sub">Biểu đồ sẽ tự động được vẽ ngay khi bạn chọn file dữ liệu tại Trang 3</div>' +
      '<button class="btn-orange-pill-sm" data-go-tab="tab-intake" type="button">Chọn file ngay ↗</button>' +
      '</div>';

    ["chartDonutPriority", "chartColumnRegion", "chartBarProducts", "chartBarStatus"].forEach(function (id) {
      var el = $(id);
      if (el) {
        el.innerHTML = emptyHtml;
        var btn = el.querySelector("[data-go-tab]");
        if (btn) btn.addEventListener("click", function () { switchPageTab("tab-intake"); });
      }
    });

    var stackedEl = $("chartStackedSalesLoad");
    if (stackedEl) {
      stackedEl.innerHTML = 
        '<div style="grid-column: 1 / -1; padding: 36px 20px; text-align: center; background: #fff5f8; border: 1.5px dashed var(--surface-pink-border); border-radius: 14px;">' +
        '<div style="font-size: 32px; margin-bottom: 8px;">👥</div>' +
        '<strong style="font-size: 14.5px; color: var(--ink);">Chưa có dữ liệu điều phối 8 NVKD</strong>' +
        '<p style="font-size: 12px; color: var(--muted); margin: 6px auto 14px; max-width: 480px;">' +
        'Biểu đồ cân bằng tải sẽ được tự động phân tích và tạo ngay khi bạn tải lên file CSV/XLSX hoặc chọn file mẫu.' +
        '</p>' +
        '<button class="btn-orange-pill" data-go-tab="tab-intake" type="button">Sang trang Xử lý Dữ liệu để chọn file 📂</button>' +
        '</div>';
      var stackedBtn = stackedEl.querySelector("[data-go-tab]");
      if (stackedBtn) stackedBtn.addEventListener("click", function () { switchPageTab("tab-intake"); });
    }

    renderLeadsTable();
    renderHitlQueue();
    renderAuditTable();
  }

  function renderDashboard() {
    if (!state.pipelineResult) {
      renderDashboardEmptyState();
      return;
    }
    renderExecutiveKpiMetrics();
    renderSvgDonutChart($("chartDonutPriority") ? "chartDonutPriority" : "chartContainerDonut");
    renderSvgRegionalChart($("chartColumnRegion") ? "chartColumnRegion" : "chartContainerRegion");
    renderSvgRankedProducts($("chartBarProducts") ? "chartBarProducts" : "chartContainerProducts");
    renderSvgPipelineFunnel($("chartBarStatus") ? "chartBarStatus" : "chartContainerPipeline");
    renderSalesMatrixGrid($("chartStackedSalesLoad") ? "chartStackedSalesLoad" : "salesMatrixContainer");
    renderLeadsTable();
    renderHitlQueue();
    renderAuditTable();
  }

  function renderExecutiveKpiMetrics() {
    var clusters = state.pipelineResult.clusters;
    var total = clusters.length;
    var tierA = clusters.filter(function (c) { return c.scoring.tier === "A"; }).length;
    var tierB = clusters.filter(function (c) { return c.scoring.tier === "B"; }).length;
    var icpCount = tierA + tierB;
    var icpPct = total > 0 ? Math.round((icpCount / total) * 100) : 0;

    var totalBudget = clusters.reduce(function (sum, c) { return sum + (c.budget_vnd || 0); }, 0);
    var pendingHitl1 = clusters.filter(function (c) {
      return c.workflow_stage === "PENDING_MANAGER_APPROVAL" || c.workflow_stage === "UNMATCHED_SALES";
    }).length;

    if ($("kpiTotalLeads")) $("kpiTotalLeads").textContent = String(total);
    if ($("kpiIcpRate")) $("kpiIcpRate").textContent = icpPct + "%";
    if ($("kpiIcpCount")) $("kpiIcpCount").textContent = icpCount + "/" + total + " DN";
    if ($("kpiTotalBudget")) $("kpiTotalBudget").textContent = formatCompactVnd(totalBudget);
    if ($("kpiDispatchedSales")) $("kpiDispatchedSales").textContent = "8/8 Sales";
    if ($("kpiPendingHitl1")) $("kpiPendingHitl1").textContent = pendingHitl1 + " Chờ duyệt TP";
    if ($("headerHitlBadge")) $("headerHitlBadge").textContent = String(pendingHitl1);
  }

  // CHART 1: SVG DONUT CHART & MULTI-METRIC TIER BREAKDOWN
  function renderSvgDonutChart(containerId) {
    var wrap = $(containerId);
    if (!wrap || !state.pipelineResult) return;
    var clusters = state.pipelineResult.clusters;
    var total = Math.max(1, clusters.length);

    var tierA = clusters.filter(function (c) { return c.scoring.tier === "A"; });
    var tierB = clusters.filter(function (c) { return c.scoring.tier === "B"; });
    var tierC = clusters.filter(function (c) { return c.scoring.tier === "C"; });
    var tierD = clusters.filter(function (c) { return c.scoring.tier === "D"; });

    var sumBudget = function (arr) { return arr.reduce(function (s, c) { return s + (c.budget_vnd || 0); }, 0); };
    var avgScore = function (arr) {
      if (arr.length === 0) return 0;
      var sum = arr.reduce(function (s, c) { return s + (c.scoring.total_score || 0); }, 0);
      return (sum / arr.length).toFixed(1);
    };

    var tierData = [
      { key: "Tier A", label: "Khẩn cấp & Chuẩn ICP", count: tierA.length, budget: sumBudget(tierA), score: avgScore(tierA), color: "#db2777" },
      { key: "Tier B", label: "Tiềm năng phát triển", count: tierB.length, budget: sumBudget(tierB), score: avgScore(tierB), color: "#ec4899" },
      { key: "Tier C", label: "Cần nuôi dưỡng thêm", count: tierC.length, budget: sumBudget(tierC), score: avgScore(tierC), color: "#c084fc" },
      { key: "Tier D", label: "Ngoài phạm vi / Sai SP", count: tierD.length, budget: sumBudget(tierD), score: avgScore(tierD), color: "#94a3b8" }
    ];

    var icpPct = Math.round(((tierA.length + tierB.length) / total) * 100);

    // SVG Donut calculation with circumferences
    var C = 2 * Math.PI * 56; // ~351.86
    var offset = 0;
    var circlesSvg = tierData.map(function (t) {
      var len = (t.count / total) * C;
      var currentOffset = offset;
      offset += len;
      return '<circle cx="80" cy="80" r="56" fill="none" stroke="' + t.color + '" stroke-width="24" stroke-dasharray="' + len.toFixed(2) + ' ' + (C - len).toFixed(2) + '" stroke-dashoffset="-' + currentOffset.toFixed(2) + '" class="chart-bar-hover" transform="rotate(-90 80 80)"><title>' + t.key + ': ' + t.count + ' DN (' + Math.round(t.count / total * 100) + '%)</title></circle>';
    }).join("");

    var svgDonut =
      '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;">' +
      '<svg viewBox="0 0 160 160" width="160" height="160" class="chart-svg-wrap">' +
      '<circle cx="80" cy="80" r="56" fill="none" stroke="#fce7f3" stroke-width="24" />' +
      circlesSvg +
      '<text x="80" y="76" text-anchor="middle" font-size="22" font-weight="900" fill="#db2777" font-family="inherit">' + icpPct + '%</text>' +
      '<text x="80" y="93" text-anchor="middle" font-size="10.5" font-weight="750" fill="#837c8e" font-family="inherit">Đạt Chuẩn ICP</text>' +
      '</svg>' +
      '</div>';

    var breakdownHtml =
      '<div style="width:100%;display:flex;flex-direction:column;gap:8px;">' +
      tierData.map(function (t) {
        var pct = Math.round((t.count / total) * 100);
        return (
          '<div style="display:flex;align-items:center;justify-content:space-between;background:#faf4f7;border:1px solid #fce7f3;border-radius:10px;padding:6px 10px;font-size:11.5px;">' +
          '<div style="display:flex;align-items:center;gap:6px;">' +
          '<span style="width:10px;height:10px;border-radius:3px;background:' + t.color + ';display:inline-block;"></span>' +
          '<strong>' + t.key + '</strong> <span style="font-size:10.5px;color:#837c8e;">(' + escHtml(t.label) + ')</span>' +
          '</div>' +
          '<div style="text-align:right;">' +
          '<b style="color:' + t.color + ';">' + t.count + ' DN (' + pct + '%)</b>' +
          '<div style="font-size:10px;color:#837c8e;">' + formatCompactVnd(t.budget) + ' • TB ' + t.score + 'đ</div>' +
          '</div>' +
          '</div>'
        );
      }).join("") +
      '</div>';

    wrap.innerHTML = svgDonut + breakdownHtml;
  }

  // CHART 2: SVG GROUPED REGIONAL BAR & REVENUE DUAL-AXIS CHART
  function renderSvgRegionalChart(containerId) {
    var wrap = $(containerId);
    if (!wrap || !state.pipelineResult) return;
    var clusters = state.pipelineResult.clusters;

    var regions = [
      { key: "Miền Bắc", label: "Bắc", count: 0, budget: 0 },
      { key: "Miền Trung", label: "Trung", count: 0, budget: 0 },
      { key: "Miền Nam", label: "Nam", count: 0, budget: 0 }
    ];

    clusters.forEach(function (c) {
      var rg = c.primary_region || "Miền Nam";
      var target = regions[2];
      if (rg.indexOf("Bắc") !== -1 || rg.indexOf("Hà Nội") !== -1) target = regions[0];
      else if (rg.indexOf("Trung") !== -1 || rg.indexOf("Đà Nẵng") !== -1) target = regions[1];
      target.count++;
      target.budget += (c.budget_vnd || 0);
    });

    var maxCount = Math.max(1, Math.max(regions[0].count, regions[1].count, regions[2].count));
    var maxBudget = Math.max(1, Math.max(regions[0].budget, regions[1].budget, regions[2].budget));
    var yStep = Math.ceil(maxCount / 4) || 5;
    var yMax = yStep * 4;

    var chartH = 130;
    var baseY = 150;

    var centers = [105, 230, 355];
    var barsSvg = regions.map(function (r, idx) {
      var cx = centers[idx];
      var hCount = Math.round((r.count / yMax) * chartH);
      var hBudget = Math.round((r.budget / maxBudget) * chartH);
      var y1 = baseY - hCount;
      var y2 = baseY - hBudget;
      var budgetStr = formatCompactVnd(r.budget);

      return (
        '<g class="chart-bar-hover">' +
        '<!-- Bar 1: Leads Count (Pink) -->' +
        '<rect x="' + (cx - 28) + '" y="' + y1 + '" width="24" height="' + hCount + '" rx="5" fill="url(#gradPinkBar)"><title>' + r.key + ': ' + r.count + ' Leads</title></rect>' +
        '<text x="' + (cx - 16) + '" y="' + (y1 - 6) + '" text-anchor="middle" font-size="11" font-weight="850" fill="#db2777" font-family="inherit">' + r.count + '</text>' +
        '<!-- Bar 2: Budget VND (Purple) -->' +
        '<rect x="' + (cx + 4) + '" y="' + y2 + '" width="24" height="' + hBudget + '" rx="5" fill="url(#gradPurpleBar)"><title>' + r.key + ': ' + budgetStr + '</title></rect>' +
        '<text x="' + (cx + 16) + '" y="' + (y2 - 6) + '" text-anchor="middle" font-size="10" font-weight="750" fill="#7e22ce" font-family="inherit">' + budgetStr + '</text>' +
        '<!-- X Label -->' +
        '<text x="' + cx + '" y="' + (baseY + 18) + '" text-anchor="middle" font-size="11.5" font-weight="850" fill="#1f1d24" font-family="inherit">' + r.key + '</text>' +
        '</g>'
      );
    }).join("");

    var gridlines = [0, 1, 2, 3, 4].map(function (step) {
      var val = yStep * step;
      var y = baseY - Math.round((val / yMax) * chartH);
      return (
        '<line x1="40" y1="' + y + '" x2="440" y2="' + y + '" class="chart-gridline" />' +
        '<text x="32" y="' + (y + 3.5) + '" text-anchor="end" class="chart-axis-text">' + val + '</text>'
      );
    }).join("");

    wrap.innerHTML =
      '<div style="display:flex;justify-content:flex-end;gap:14px;font-size:11px;font-weight:750;margin-bottom:6px;">' +
      '<span style="display:flex;align-items:center;gap:5px;color:#db2777;"><i style="width:10px;height:10px;background:#db2777;border-radius:3px;display:inline-block;"></i> Số Leads (DN)</span>' +
      '<span style="display:flex;align-items:center;gap:5px;color:#7e22ce;"><i style="width:10px;height:10px;background:#7e22ce;border-radius:3px;display:inline-block;"></i> Doanh số ước tính</span>' +
      '</div>' +
      '<svg viewBox="0 0 460 185" width="100%" height="185" class="chart-svg-wrap">' +
      '<defs>' +
      '<linearGradient id="gradPinkBar" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#ec4899"/><stop offset="100%" stop-color="#db2777"/></linearGradient>' +
      '<linearGradient id="gradPurpleBar" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#a855f7"/><stop offset="100%" stop-color="#7e22ce"/></linearGradient>' +
      '</defs>' +
      gridlines +
      barsSvg +
      '<line x1="40" y1="' + baseY + '" x2="440" y2="' + baseY + '" stroke="#ebdce4" stroke-width="1.5" />' +
      '</svg>';
  }

  // CHART 3: RANKED HORIZONTAL PRODUCTS DEMAND & REVENUE
  function renderSvgRankedProducts(containerId) {
    var wrap = $(containerId);
    if (!wrap || !state.pipelineResult) return;
    var clusters = state.pipelineResult.clusters;

    var prodMap = {};
    clusters.forEach(function (c) {
      var p = c.analysis.primary_product_code || "Khác";
      if (!prodMap[p]) prodMap[p] = { code: p, count: 0, budget: 0 };
      prodMap[p].count++;
      prodMap[p].budget += (c.budget_vnd || 0);
    });

    var sorted = Object.keys(prodMap).map(function (k) { return prodMap[k]; })
      .sort(function (a, b) { return b.count - a.count; })
      .slice(0, 5);

    var maxCount = Math.max(1, sorted[0] ? sorted[0].count : 1);
    var totalClusters = Math.max(1, clusters.length);

    wrap.innerHTML =
      '<div class="ranked-products-list">' +
      sorted.map(function (p, idx) {
        var pct = Math.round((p.count / totalClusters) * 100);
        var barW = Math.round((p.count / maxCount) * 100);
        var budgetStr = formatCompactVnd(p.budget);

        return (
          '<div class="ranked-product-row">' +
          '<div class="ranked-product-header">' +
          '<div>' +
          '<span class="ranked-prod-badge">#' + (idx + 1) + '</span>' +
          '<strong>' + escHtml(p.code) + '</strong>' +
          '</div>' +
          '<div style="font-weight:800;color:#db2777;">' +
          p.count + ' DN • <span style="color:#7e22ce;">' + budgetStr + '</span> <span style="font-size:10.5px;color:#837c8e;">(' + pct + '%)</span>' +
          '</div>' +
          '</div>' +
          '<div class="ranked-bar-track">' +
          '<div class="ranked-bar-fill" style="width:' + barW + '%;"></div>' +
          '</div>' +
          '</div>'
        );
      }).join("") +
      '</div>';
  }

  // CHART 4: PIPELINE FUNNEL & SLA BOTTLENECK TRACKER
  function renderSvgPipelineFunnel(containerId) {
    var wrap = $(containerId);
    if (!wrap || !state.pipelineResult) return;
    var clusters = state.pipelineResult.clusters;
    var total = Math.max(1, clusters.length);

    var assigned = clusters.filter(function (c) { return c.workflow_stage === "ASSIGNED_TO_REP"; }).length;
    var pending = clusters.filter(function (c) { return c.workflow_stage === "PENDING_MANAGER_APPROVAL" || c.workflow_stage === "UNMATCHED_SALES"; }).length;
    var readyHitl2 = clusters.filter(function (c) { return c.next_action && c.next_action.draft_channel; }).length;
    var tierAB = clusters.filter(function (c) { return c.scoring.tier === "A" || c.scoring.tier === "B"; }).length;

    var stages = [
      { step: 1, title: "Tiếp nhận & Làm sạch dữ liệu", count: total, pct: 100, note: "0 lỗi định dạng" },
      { step: 2, title: "Lọc trùng lặp & Đạt chuẩn ICP", count: tierAB, pct: Math.round(tierAB / total * 100), note: "Tier A & B vượt chuẩn" },
      { step: 3, title: "Gán việc tự động cho 8 NVKD", count: assigned, pct: Math.round(assigned / total * 100), note: "Khớp địa bàn & tải tối ưu" },
      { step: 4, title: "Chốt chặn Trưởng phòng [HITL 1]", count: pending, pct: Math.round(pending / total * 100), note: pending > 0 ? "Cần phê duyệt ngoại lệ" : "Không có tắc nghẽn" },
      { step: 5, title: "NVKD sẵn sàng tiếp cận [HITL 2]", count: readyHitl2, pct: Math.round(readyHitl2 / total * 100), note: "Đã tạo nháp kịch bản" }
    ];

    wrap.innerHTML =
      '<div class="pipeline-funnel-stack">' +
      stages.map(function (s) {
        var isPending = s.step === 4 && s.count > 0;
        var tagColor = isPending ? "#b45309" : "#047857";
        var tagBg = isPending ? "#fef3c7" : "#ecfdf5";

        return (
          '<div class="pipeline-stage-item">' +
          '<div class="pipeline-stage-left">' +
          '<span class="pipeline-num-circle">' + s.step + '</span>' +
          '<div>' +
          '<div class="pipeline-stage-title">' + escHtml(s.title) + '</div>' +
          '<div class="pipeline-stage-meta">' + escHtml(s.note) + '</div>' +
          '</div>' +
          '</div>' +
          '<div class="pipeline-stage-right">' +
          '<span>' + s.count + ' DN (' + s.pct + '%)</span>' +
          '<span style="background:' + tagBg + ';color:' + tagColor + ';font-size:10px;padding:2px 7px;border-radius:6px;font-weight:800;">✓ Hoàn thành</span>' +
          '</div>' +
          '</div>'
        );
      }).join("") +
      '</div>';
  }

  // CHART 5: 8 SALES REPS WORKLOAD & CAPACITY MATRIX
  function renderSalesMatrixGrid(containerId) {
    var wrap = $(containerId);
    if (!wrap || !state.pipelineResult) return;

    var reps = NAMVIET_DATA.salesTeam.sales_reps;
    var wMap = {};
    state.pipelineResult.insights.rep_workloads.forEach(function (w) { wMap[w.id] = w; });

    wrap.innerHTML = reps.map(function (r) {
      var w = wMap[r.id] || { existing_leads: r.current_active_leads, new_assigned: 0, capacity: r.weekly_capacity };
      var total = w.existing_leads + w.new_assigned;
      var cap = r.weekly_capacity || 30;
      var pct = Math.min(100, Math.round((total / cap) * 100));

      var existPct = Math.round((w.existing_leads / cap) * 100);
      var newPct = Math.round((w.new_assigned / cap) * 100);

      var statusText = "🟢 Tối ưu";
      var statusColor = "#047857";
      var statusBg = "#ecfdf5";
      if (pct >= 90) {
        statusText = "🔴 Quá tải";
        statusColor = "#b91c1c";
        statusBg = "#fef2f2";
      } else if (pct >= 70) {
        statusText = "🟡 Cận tải";
        statusColor = "#b45309";
        statusBg = "#fef3c7";
      }

      var initials = r.name.split(" ").map(function (n) { return n[0]; }).slice(-2).join("");

      return (
        '<div class="sales-rep-card">' +
        '<div class="sales-rep-top">' +
        '<div class="sales-rep-avatar">' + escHtml(initials) + '</div>' +
        '<div class="sales-rep-info">' +
        '<strong>' + escHtml(r.name) + '</strong>' +
        '<small>' + escHtml(r.region) + ' • ' + (r.specialty_products || []).join(", ") + '</small>' +
        '</div>' +
        '</div>' +
        '<div class="sales-load-metric-row">' +
        '<span>Tổng tải: <b>' + total + '/' + cap + '</b></span>' +
        '<span style="color:#db2777;">' + pct + '% công suất</span>' +
        '</div>' +
        '<div class="sales-rep-progress-track">' +
        '<div class="seg-existing" style="width:' + existPct + '%;" title="Hiện có: ' + w.existing_leads + '"></div>' +
        '<div class="seg-new" style="width:' + newPct + '%;" title="Giao mới: ' + w.new_assigned + '"></div>' +
        '</div>' +
        '<div class="sales-rep-badge-row">' +
        '<span style="background:' + statusBg + ';color:' + statusColor + ';padding:2px 7px;border-radius:6px;font-weight:800;">' + statusText + '</span>' +
        '<button class="btn-preset-pill-sm btn-filter-rep" data-rep="' + escHtml(r.name) + '" type="button">Xem Leads ↗</button>' +
        '</div>' +
        '</div>'
      );
    }).join("");

    wrap.querySelectorAll(".btn-filter-rep").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var repName = btn.dataset.rep;
        switchPageTab("tab-leads");
        if ($("searchInput")) {
          $("searchInput").value = repName;
          state.searchQuery = repName;
          renderLeadsTable();
        }
      });
    });
  }

  function renderLeadsTable() {
    var table = $("clusterTable");
    if (!table) return;
    var tbody = table.querySelector("tbody");
    if (!tbody) return;

    if (!state.pipelineResult) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:36px;color:#837c8e;">Chưa có dữ liệu khách hàng. Vui lòng chọn hoặc nạp file tại Trang 3 [Xử lý Dữ liệu] để AI Agent phân tích và hiển thị danh sách.</td></tr>';
      return;
    }

    var list = state.pipelineResult.clusters.filter(function (c) {
      if (state.filterStage !== "ALL" && c.workflow_stage !== state.filterStage) return false;
      if (state.filterTier !== "ALL" && c.scoring.tier !== state.filterTier) return false;
      if (state.searchQuery) {
        var q = state.searchQuery.toLowerCase();
        var txt = (c.company_name + " " + c.tax_id + " " + c.primary_region + " " + (c.routing.proposed_rep ? c.routing.proposed_rep.name : "")).toLowerCase();
        if (txt.indexOf(q) === -1) return false;
      }
      return true;
    });

    if (list.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:24px;color:#a8a29e;">Không có doanh nghiệp nào khớp bộ lọc.</td></tr>';
      return;
    }

    tbody.innerHTML = list.map(function (c, idx) {
      var repName = c.routing.proposed_rep ? c.routing.proposed_rep.name : "Chưa giao";
      return (
        '<tr data-cid="' + escHtml(c.cluster_id) + '">' +
        '<td>' + (idx + 1) + '</td>' +
        '<td><strong>' + escHtml(c.company_name) + '</strong><div style="font-size:10.5px;color:#78716c;">' + (c.aliases.join(", ") || c.sources.join(", ")) + '</div></td>' +
        '<td>' + escHtml(c.analysis.primary_product_code) + '<div style="font-size:10.5px;color:#78716c;">' + escHtml(c.primary_region || "") + '</div></td>' +
        '<td><b>Hạng ' + c.scoring.tier + '</b> (' + c.scoring.total_score + 'đ)</td>' +
        '<td>' + escHtml(formatCompactVnd(c.budget_vnd)) + '</td>' +
        '<td>' + escHtml(repName) + '</td>' +
        '<td><span style="background:#fefce8;border:1px solid #fef08a;color:#854d0e;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:750;">' + escHtml(c.workflow_stage_label || c.workflow_stage) + '</span></td>' +
        '<td><button class="btn-preset-pill-sm btn-inspect-row" data-cid="' + escHtml(c.cluster_id) + '" type="button">Chi tiết ▾</button></td>' +
        '</tr>'
      );
    }).join("");

    tbody.querySelectorAll("tr[data-cid]").forEach(function (tr) {
      tr.addEventListener("click", function () {
        openInspector(tr.dataset.cid);
      });
    });
  }

  function renderHitlQueue() {
    var box = $("hitlQueueContainer");
    if (!box) return;

    if (!state.pipelineResult) {
      if ($("subtabHitlCount")) $("subtabHitlCount").textContent = "0";
      box.innerHTML = '<div style="background:#fff5f8;border:1.5px dashed var(--surface-pink-border);border-radius:14px;padding:32px;text-align:center;color:#837c8e;">Chưa có công việc chờ duyệt. Hãy chọn file dữ liệu ở Bước 3 để bắt đầu.</div>';
      return;
    }

    var pending = state.pipelineResult.clusters.filter(function (c) {
      return c.workflow_stage === "PENDING_MANAGER_APPROVAL" || c.workflow_stage === "UNMATCHED_SALES";
    });

    if ($("subtabHitlCount")) $("subtabHitlCount").textContent = String(pending.length);

    if (pending.length === 0) {
      box.innerHTML = '<div style="background:#fff;border:1px solid #e7e2da;border-radius:14px;padding:24px;text-align:center;color:#047857;font-weight:750;">✓ Tuyệt vời! Hiện không có hồ sơ nào bị tắc nghẽn hoặc cần phê duyệt ngoại lệ.</div>';
      return;
    }

    box.innerHTML = pending.map(function (c) {
      return (
        '<div style="background:#fff;border:1.5px solid #fef08a;border-radius:14px;padding:16px;margin-bottom:12px;">' +
        '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px;">' +
        '<div><strong style="font-size:14px;">' + escHtml(c.company_name) + '</strong> (' + escHtml(c.cluster_id) + ')</div>' +
        '<span style="background:#fef2f2;color:#b91c1c;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:800;">Chờ Trưởng phòng duyệt [HITL 1]</span>' +
        '</div>' +
        '<p style="font-size:12px;color:#78716c;margin-bottom:10px;">Lý do ngoại lệ: ' + escHtml(c.routing.routing_reason) + '</p>' +
        '<div style="display:flex;gap:8px;">' +
        '<button class="btn-orange-pill" style="font-size:11.5px;padding:6px 14px;" onclick="openInspector(\'' + c.cluster_id + '\')">Mở Duyệt &amp; Giao Việc ↗</button>' +
        '</div></div>'
      );
    }).join("");
  }

  function renderAuditTable() {
    var table = $("auditFullTable");
    if (!table) return;
    var tbody = table.querySelector("tbody");
    if (!tbody) return;

    if (!state.pipelineResult) {
      tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:32px;color:#837c8e;">Chưa có log truy vết. Hãy chọn file để bắt đầu xử lý.</td></tr>';
      return;
    }

    var logs = [];
    state.pipelineResult.clusters.forEach(function (c) {
      (c.audit_logs || []).forEach(function (l) {
        logs.push({ id: c.cluster_id, name: c.company_name, step: l.step, detail: l.detail });
      });
    });

    tbody.innerHTML = logs.slice(0, 50).map(function (lg) {
      return '<tr>' +
        '<td><b>' + escHtml(lg.id) + '</b></td>' +
        '<td>' + escHtml(lg.name) + '</td>' +
        '<td><strong>' + escHtml(lg.step) + '</strong></td>' +
        '<td style="white-space:normal;">' + escHtml(lg.detail) + '</td>' +
        '</tr>';
    }).join("");
  }

  function openInspector(clusterId) {
    state.selectedClusterId = clusterId;
    var pane = $("inspectorPane");
    if (!pane || !state.pipelineResult) return;

    var c = state.pipelineResult.clusters.filter(function (item) { return item.cluster_id === clusterId; })[0];
    if (!c) return;

    if ($("drawerTitle")) $("drawerTitle").textContent = c.company_name;
    if ($("drawerClusterMeta")) $("drawerClusterMeta").textContent = c.cluster_id + " • " + c.sources.join(", ");

    var repOpts = NAMVIET_DATA.salesTeam.sales_reps.map(function (r) {
      var sel = c.routing.proposed_rep && c.routing.proposed_rep.id === r.id ? " selected" : "";
      return '<option value="' + escHtml(r.id) + '"' + sel + ">" + escHtml(r.name + " (" + r.region + ")") + "</option>";
    }).join("");

    pane.innerHTML =
      '<div style="margin-bottom:16px;">' +
      '<div style="font-size:12px;color:#78716c;margin-bottom:4px;">Nhu cầu: <strong>' + escHtml(c.analysis.primary_product_name) + '</strong></div>' +
      '<div style="font-size:12px;color:#78716c;margin-bottom:10px;">Ngân sách: <strong>' + escHtml(formatCompactVnd(c.budget_vnd)) + '</strong> • Điểm: <strong>' + c.scoring.total_score + 'đ (Hạng ' + c.scoring.tier + ')</strong></div>' +
      '<div style="background:#fcfbf9;border:1px solid #f0ebe4;padding:12px;border-radius:10px;font-size:12.5px;margin-bottom:14px;">' + escHtml(c.combined_need) + '</div>' +

      '<!-- HITL 1 BOX -->' +
      '<div style="background:#fffbeb;border:1.5px solid #fef08a;border-radius:12px;padding:14px;margin-bottom:14px;">' +
      '<strong style="color:#b45309;display:block;margin-bottom:6px;">[HITL 1] Quyết Định Của Trưởng Phòng KD</strong>' +
      '<div style="margin-bottom:8px;"><label style="font-size:11.5px;font-weight:700;">Giao NVKD phụ trách:</label><br>' +
      '<select id="mgrRepSelect" style="width:100%;padding:6px;border-radius:6px;border:1px solid #e7e2da;margin-top:2px;">' + repOpts + '</select></div>' +
      '<div style="display:flex;gap:6px;">' +
      '<button id="btnDrawerApprove" class="btn-orange-pill" type="button">✓ Duyệt &amp; Giao Việc</button>' +
      '<button id="btnDrawerHold" class="btn-outline-pill" type="button">Tạm giữ</button>' +
      '</div></div>' +

      '<!-- HITL 2 BOX -->' +
      '<div style="background:#f0fdf4;border:1.5px solid #bbf7d0;border-radius:12px;padding:14px;">' +
      '<strong style="color:#166534;display:block;margin-bottom:6px;">[HITL 2] Sales Duyệt Bản Nháp Tiếp Cận</strong>' +
      '<textarea id="drawerDraft" rows="6" style="width:100%;padding:8px;border-radius:8px;border:1px solid #bbf7d0;font-size:12px;box-sizing:border-box;">' + escHtml(c.draft ? c.draft.message : "") + '</textarea>' +
      '<div style="display:flex;gap:6px;margin-top:8px;">' +
      '<button id="btnDrawerCopy" class="btn-outline-pill" type="button">Sao chép</button>' +
      '<button id="btnDrawerSend" class="btn-orange-pill" type="button">✓ Xác nhận đã duyệt gửi</button>' +
      '</div></div>' +
      '</div>';

    if ($("inspectorModal")) $("inspectorModal").classList.remove("hidden");

    if ($("btnDrawerApprove")) {
      $("btnDrawerApprove").addEventListener("click", function () {
        NamVietAgent.applyManagerDecision(c, "APPROVE", $("mgrRepSelect").value, "Trưởng phòng duyệt", NAMVIET_DATA.salesTeam);
        renderDashboard();
        openInspector(clusterId);
        showToast("Đã duyệt phân công cho NVKD!", "success");
      });
    }
    if ($("btnDrawerHold")) {
      $("btnDrawerHold").addEventListener("click", function () {
        NamVietAgent.applyManagerDecision(c, "REJECT_HOLD", null, "Tạm giữ rà soát", NAMVIET_DATA.salesTeam);
        renderDashboard();
        openInspector(clusterId);
        showToast("Đã tạm giữ hồ sơ rà soát.", "success");
      });
    }
    if ($("btnDrawerCopy")) {
      $("btnDrawerCopy").addEventListener("click", function () {
        var text = $("drawerDraft") ? $("drawerDraft").value : "";
        if (navigator.clipboard) {
          navigator.clipboard.writeText(text).then(function () {
            showToast("Đã sao chép nội dung nháp!", "success");
          });
        }
      });
    }
    if ($("btnDrawerSend")) {
      $("btnDrawerSend").addEventListener("click", function () {
        NamVietAgent.applyRepReviewDecision(c, "Email + Zalo", $("drawerDraft") ? $("drawerDraft").value : "", "");
        renderDashboard();
        openInspector(clusterId);
        showToast("Đã xác nhận kiểm duyệt gửi tin!", "success");
      });
    }
  }

  function closeInspector() {
    if ($("inspectorModal")) $("inspectorModal").classList.add("hidden");
  }

  // =========================================================================
  // 5. Exports (Excel, CSV, Standalone HTML)
  // =========================================================================
  function triggerDownload(blob, filename) {
    if (typeof URL === "undefined" || !URL.createObjectURL) return;
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      if (a.parentNode) a.parentNode.removeChild(a);
      URL.revokeObjectURL(url);
    }, 100);
  }

  function exportResultsCsv() {
    if (!state.pipelineResult) return;
    var headers = ["Mã Cụm", "Doanh nghiệp", "Khu vực", "Sản phẩm", "Hạng", "Điểm", "Ngân sách", "NVKD", "Trạng thái"];
    var rows = state.pipelineResult.clusters.map(function (c) {
      return [
        c.cluster_id,
        c.company_name,
        c.primary_region,
        c.analysis.primary_product_code,
        c.scoring.tier,
        c.scoring.total_score,
        c.budget_vnd || 0,
        c.routing.proposed_rep ? c.routing.proposed_rep.name : "",
        c.workflow_stage_label || c.workflow_stage
      ];
    });
    var csv = SpreadsheetIO.toCSV(headers, rows);
    triggerDownload(new Blob([csv], { type: "text/csv;charset=utf-8;" }), "NamViet_Leads_Report.csv");
  }

  function exportResultsXlsx() {
    if (!state.pipelineResult) return;
    var headers = ["Mã Cụm", "Doanh nghiệp", "Khu vực", "Sản phẩm", "Hạng", "Điểm", "Ngân sách", "NVKD", "Trạng thái"];
    var rows = state.pipelineResult.clusters.map(function (c) {
      return [
        c.cluster_id,
        c.company_name,
        c.primary_region,
        c.analysis.primary_product_code,
        c.scoring.tier,
        c.scoring.total_score,
        c.budget_vnd || 0,
        c.routing.proposed_rep ? c.routing.proposed_rep.name : "",
        c.workflow_stage_label || c.workflow_stage
      ];
    });
    var bytes = SpreadsheetIO.buildXlsxUint8Array([{ name: "Leads_NamViet", headers: headers, rows: rows }]);
    triggerDownload(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "NamViet_Leads_Report.xlsx");
  }

  // =========================================================================
  // 6. Application Initialization
  // =========================================================================
  function init() {
    // 4 Top Tabs
    document.querySelectorAll(".nav-tab-pill").forEach(function (pill) {
      pill.addEventListener("click", function () {
        switchPageTab(pill.dataset.tab);
      });
    });

    if ($("brandLogoBtn")) {
      $("brandLogoBtn").addEventListener("click", function () { switchPageTab("tab-dashboard"); });
    }

    // Top Header Action Buttons
    if ($("btnHeaderQuickLoad50")) {
      $("btnHeaderQuickLoad50").addEventListener("click", loadSampleSalesReport50);
    }
    if ($("btnDashReload50")) {
      $("btnDashReload50").addEventListener("click", loadSampleSalesReport50);
    }
    if ($("btnHeaderExportHtml")) {
      $("btnHeaderExportHtml").addEventListener("click", exportResultsCsv);
    }

    // Navigation jumps across all tabs
    document.querySelectorAll("[data-go-tab], [data-jump-tab]").forEach(function (el) {
      el.addEventListener("click", function () {
        var target = el.dataset.goTab || el.dataset.jumpTab;
        if (target) switchPageTab(target);
      });
    });

    // Step rows on Page 2 (Workflow)
    document.querySelectorAll(".wf-step-row-item").forEach(function (row) {
      row.addEventListener("mouseenter", function () {
        updateStepPreview(parseInt(row.dataset.step, 10));
      });
      row.addEventListener("click", function () {
        updateStepPreview(parseInt(row.dataset.step, 10));
      });
    });

    // Ribbon nodes on Page 2 (Workflow)
    document.querySelectorAll(".wf-ribbon-node").forEach(function (node) {
      node.addEventListener("click", function () {
        updateStepPreview(parseInt(node.dataset.step, 10));
      });
    });

    // Preset buttons on Intake Studio
    if ($("btnPresetStandard")) {
      $("btnPresetStandard").addEventListener("click", function () {
        loadDefaultDemoData();
        runAgentNow({ openDashboard: true });
        showToast("Đã nạp 18 leads mẫu chuẩn!", "success");
      });
    }
    if ($("btnPresetError")) {
      $("btnPresetError").addEventListener("click", loadSampleDirtyLeads);
    }
    if ($("btnPreset50Sales")) {
      $("btnPreset50Sales").addEventListener("click", loadSampleSalesReport50);
    }
    if ($("btnPreset700")) {
      $("btnPreset700").addEventListener("click", loadSimulatedLargeBatch);
    }
    if ($("btnIntakeLoad50")) $("btnIntakeLoad50").addEventListener("click", loadSampleSalesReport50);

    // 4 Intake Mode Switcher Pills (Ảnh 1)
    document.querySelectorAll(".intake-mode-pill").forEach(function (pill) {
      pill.addEventListener("click", function () {
        var mode = pill.dataset.intakeMode;
        document.querySelectorAll(".intake-mode-pill").forEach(function (p) {
          p.classList.toggle("active", p === pill);
        });
        if ($("panelIntakeCsv")) $("panelIntakeCsv").classList.toggle("hidden", mode !== "csv");
        if ($("panelIntakeGsheet")) $("panelIntakeGsheet").classList.toggle("hidden", mode !== "gsheet");
        if ($("panelIntakePaste")) $("panelIntakePaste").classList.toggle("hidden", mode !== "paste");
        if ($("panelIntakeQuick")) $("panelIntakeQuick").classList.toggle("hidden", mode !== "quick");
      });
    });

    // Parse pasted table from Excel / Sheets
    if ($("btnParsePastedTable")) {
      $("btnParsePastedTable").addEventListener("click", function () {
        var text = ($("pasteTableInput") ? $("pasteTableInput").value : "").trim();
        if (!text) {
          showStatusBanner("Vui lòng dán dữ liệu bảng từ Excel/Sheets vào khung trên.", "error");
          return;
        }
        var tbl = SpreadsheetIO.parseCSV(text);
        if (!tbl || !tbl.records || tbl.records.length === 0) {
          showStatusBanner("Không nhận diện được dòng dữ liệu hợp lệ trong nội dung dán.", "error");
          return;
        }
        loadTableIntoIntake(tbl.headers, tbl.records, "Bảng dán từ Clipboard");
        showStatusBanner("Đã nạp thành công " + tbl.records.length + " dòng từ bảng dán.", "success");
        var autoRun = $("chkAutoRunOnSelect") ? $("chkAutoRunOnSelect").checked : true;
        if (autoRun) runAgentNow({ openDashboard: true });
      });
    }

    // Quick Single Lead test
    if ($("btnRunQuickSingleLead")) {
      $("btnRunQuickSingleLead").addEventListener("click", function () {
        var company = ($("quickCompanyName") ? $("quickCompanyName").value : "").trim();
        var need = ($("quickNeed") ? $("quickNeed").value : "").trim();
        var region = ($("quickRegion") ? $("quickRegion").value : "Hà Nội").trim();
        var budgetRaw = ($("quickBudget") ? $("quickBudget").value : "50,000,000").trim();
        if (!company) {
          showStatusBanner("Vui lòng nhập tên doanh nghiệp.", "error");
          return;
        }
        var row = {
          "MÃ DÒNG": "DS-TEST-01",
          "THÁNG": "2026-11",
          "KÊNH BÁN": "Website Trực tiếp",
          "TÊN DOANH NGHIỆP": company,
          "NHU CẦU": need || "Tư vấn phần mềm quản trị",
          "KHU VỰC": region,
          "NGÂN SÁCH (VND)": budgetRaw,
          "SẢN PHẨM": need,
          "SỐ ĐƠN": "1"
        };
        var headers = Object.keys(row);
        loadTableIntoIntake(headers, [row], "Lead nhập nhanh đơn lẻ");
        runAgentNow({ openDashboard: true });
        showToast("Đã phân tích lead \"" + company + "\" thành công!", "success");
      });
    }

    // Downloads
    if ($("btnPresetDownload")) {
      $("btnPresetDownload").addEventListener("click", function () {
        window.location.href = "Du_Lieu_Mau_NamViet_Test.xlsx";
      });
    }
    if ($("btnDownload50Csv")) {
      $("btnDownload50Csv").addEventListener("click", function () {
        var raw = NAMVIET_DATA.sampleSalesReport50 || [];
        if (raw.length === 0) return;
        var headers = Object.keys(raw[0]);
        var rows = raw.map(function (r) { return headers.map(function (h) { return r[h] == null ? "" : r[h]; }); });
        triggerDownload(new Blob([SpreadsheetIO.toCSV(headers, rows)], { type: "text/csv;charset=utf-8;" }), "BaoCao_50Dong.csv");
      });
    }
    if ($("btnDownloadSample18Csv")) {
      $("btnDownloadSample18Csv").addEventListener("click", function () {
        var rows = NAMVIET_DATA.sampleLeads.map(function (r) { return DEFAULT_HEADERS.map(function (h) { return r[h] == null ? "" : r[h]; }); });
        triggerDownload(new Blob([SpreadsheetIO.toCSV(DEFAULT_HEADERS, rows)], { type: "text/csv;charset=utf-8;" }), "NamViet_18Leads.csv");
      });
    }
    if ($("btnDownloadSample18Xlsx")) {
      $("btnDownloadSample18Xlsx").addEventListener("click", function () {
        window.location.href = "Du_Lieu_Mau_NamViet_Test.xlsx";
      });
    }

    // Drag and Drop
    var dropZone = $("mainDropZone");
    var fileInput = $("mainFileInput");
    if (dropZone && fileInput) {
      dropZone.addEventListener("dragover", function (e) { e.preventDefault(); dropZone.classList.add("dragover"); });
      dropZone.addEventListener("dragleave", function () { dropZone.classList.remove("dragover"); });
      dropZone.addEventListener("drop", function (e) {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
          var autoRun = $("chkAutoRunOnSelect") ? $("chkAutoRunOnSelect").checked : true;
          handleUploadedFile(e.dataTransfer.files[0], autoRun);
        }
      });
      fileInput.addEventListener("change", function (e) {
        if (e.target.files && e.target.files[0]) {
          var autoRun = $("chkAutoRunOnSelect") ? $("chkAutoRunOnSelect").checked : true;
          handleUploadedFile(e.target.files[0], autoRun);
          e.target.value = "";
        }
      });
    }

    // Hero & Intro buttons
    if ($("btnHeroStart")) {
      $("btnHeroStart").addEventListener("click", function () { switchPageTab("tab-intake"); });
    }
    if ($("btnHeroArrow")) {
      $("btnHeroArrow").addEventListener("click", function () { switchPageTab("tab-intake"); });
    }
    if ($("btnIntroLoad50")) {
      $("btnIntroLoad50").addEventListener("click", function () { loadSampleSalesReport50(); });
    }
    if ($("btnIntroGoIntake")) {
      $("btnIntroGoIntake").addEventListener("click", function () { switchPageTab("tab-intake"); });
    }

    // Toggle mapping detail
    if ($("btnToggleMapping")) {
      $("btnToggleMapping").addEventListener("click", function () {
        if ($("mappingGrid")) $("mappingGrid").classList.toggle("hidden");
      });
    }

    // Google Sheets fetch
    if ($("btnFetchGsheet")) {
      $("btnFetchGsheet").addEventListener("click", function () {
        var input = $("gsheetUrlInput");
        var url = (input ? input.value : "").trim();
        if (!url) {
          showStatusBanner("Vui lòng nhập link Google Sheet công khai.", "error");
          return;
        }
        showStatusBanner("Đang tải dữ liệu từ Google Sheet...", "success");
        SpreadsheetIO.fetchGoogleSheetTable(url)
          .then(function (res) {
            loadTableIntoIntake(res.table.headers, res.table.records, "Google Sheet");
            showStatusBanner("Đã nạp thành công Google Sheet (" + res.table.records.length + " dòng).", "success");
            var autoRun = $("chkAutoRunOnSelect") ? $("chkAutoRunOnSelect").checked : true;
            if (autoRun) runAgentNow({ openDashboard: true });
          })
          .catch(function (err) {
            showStatusBanner(err.message, "error");
          });
      });
    }

    // Tab 4 Subtabs (.dash-sub-pill)
    document.querySelectorAll(".dash-sub-pill").forEach(function (pill) {
      pill.addEventListener("click", function () {
        switchDashSubTab(pill.dataset.subtab);
      });
    });

    // Run Agent
    if ($("btnRunAgent")) {
      $("btnRunAgent").addEventListener("click", function () {
        runAgentNow({ openDashboard: true });
      });
    }

    // Exports
    if ($("btnExportLeadsCsv")) $("btnExportLeadsCsv").addEventListener("click", exportResultsCsv);
    if ($("btnExportTasksCsv")) $("btnExportTasksCsv").addEventListener("click", exportResultsCsv);
    if ($("btnExportJsonl")) $("btnExportJsonl").addEventListener("click", exportResultsCsv);
    if ($("btnExportXlsxFull")) $("btnExportXlsxFull").addEventListener("click", exportResultsXlsx);
    if ($("btnFinalizeAndExportHtml")) {
      $("btnFinalizeAndExportHtml").addEventListener("click", function () {
        if ($("dashReviewStatusTag")) {
          $("dashReviewStatusTag").textContent = "Đã chốt kết quả ✓";
          $("dashReviewStatusTag").style.background = "#dcfce7";
          $("dashReviewStatusTag").style.color = "#15803d";
          $("dashReviewStatusTag").style.borderColor = "#86efac";
        }
        exportResultsCsv();
        showToast("Đã chốt kết quả thành công và xuất dữ liệu!", "success");
      });
    }

    // Search and filter in Leads table
    if ($("searchInput")) {
      $("searchInput").addEventListener("input", function () {
        state.searchQuery = $("searchInput").value.trim();
        renderLeadsTable();
      });
    }
    if ($("filterStage")) {
      $("filterStage").addEventListener("change", function () {
        state.filterStage = $("filterStage").value;
        renderLeadsTable();
      });
    }
    if ($("filterTier")) {
      $("filterTier").addEventListener("change", function () {
        state.filterTier = $("filterTier").value;
        renderLeadsTable();
      });
    }

    // Inspector
    if ($("btnCloseInspector")) $("btnCloseInspector").addEventListener("click", closeInspector);
    if ($("inspectorBackdrop")) $("inspectorBackdrop").addEventListener("click", closeInspector);

    // Feedback System
    initFeedbackSystem();

    // Initial boot: Start with empty dashboard; dashboard is only created when a file is chosen!
    renderDashboard();

    // Default to Tab 1 (1. Giới thiệu) exactly as requested
    switchPageTab("tab-intro");
  }

  // =========================================================================
  // FEEDBACK SYSTEM (Góp ý & Đánh giá)
  // =========================================================================
  var currentSelectedStars = 5;

  function initFeedbackSystem() {
    var starLabels = {
      1: "1/5 Cần cải thiện",
      2: "2/5 Tạm chấp nhận",
      3: "3/5 Đạt yêu cầu",
      4: "4/5 Tốt & Hiệu quả",
      5: "5/5 Rất xuất sắc"
    };

    // Star rating picker
    document.querySelectorAll(".star-rating-picker .star-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var stars = parseInt(btn.dataset.stars, 10);
        currentSelectedStars = stars;
        document.querySelectorAll(".star-rating-picker .star-btn").forEach(function (b) {
          var s = parseInt(b.dataset.stars, 10);
          b.classList.toggle("active", s <= stars);
        });
        if ($("starRatingLabel")) $("starRatingLabel").textContent = starLabels[stars] || (stars + "/5");
      });
    });

    // Scroll to feedback section helper
    function scrollToFeedbackSection() {
      switchPageTab("tab-feedback");
      setTimeout(function () {
        var el = $("sectionCustomerFeedback");
        if (el) {
          el.scrollIntoView({ behavior: "smooth", block: "start" });
          if ($("feedbackUserName")) $("feedbackUserName").focus();
        }
      }, 100);
    }

    if ($("btnFloatingFeedback")) {
      $("btnFloatingFeedback").addEventListener("click", scrollToFeedbackSection);
    }
    if ($("btnDashFeedbackShortcut")) {
      $("btnDashFeedbackShortcut").addEventListener("click", scrollToFeedbackSection);
    }

    // Submit Feedback
    if ($("feedbackFormSubmit")) {
      $("feedbackFormSubmit").addEventListener("submit", function (e) {
        e.preventDefault();
        var nameInput = $("feedbackUserName");
        var compInput = $("feedbackUserCompany");
        var catSelect = $("feedbackCategory");
        var commInput = $("feedbackUserComment");

        var name = (nameInput && nameInput.value.trim()) || "Người dùng ẩn danh";
        var company = (compInput && compInput.value.trim()) || "Doanh nghiệp đối tác";
        var comment = (commInput && commInput.value.trim()) || "";
        var category = catSelect ? catSelect.options[catSelect.selectedIndex].text : "Đóng góp chung";

        if (!comment) {
          showToast("Vui lòng nhập nội dung phản hồi!", "error");
          return;
        }

        var newFeedback = {
          name: name,
          company: company,
          stars: currentSelectedStars,
          comment: comment,
          category: category,
          created_at: new Date().toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" }) + " hôm nay"
        };

        // Save to localStorage
        try {
          var list = JSON.parse(localStorage.getItem("namviet_feedbacks") || "[]");
          list.unshift(newFeedback);
          localStorage.setItem("namviet_feedbacks", JSON.stringify(list));
        } catch (err) {}

        // Add to cards grid visually
        appendFeedbackCardToGrid(newFeedback);

        // Reset form
        if (nameInput) nameInput.value = "";
        if (compInput) compInput.value = "";
        if (commInput) commInput.value = "";

        showToast("✓ Cảm ơn bạn! Phản hồi (" + currentSelectedStars + "★) đã được ghi nhận thành công.", "success");
      });
    }

    // Load persisted feedbacks
    loadPersistedFeedbacks();
  }

  function appendFeedbackCardToGrid(item) {
    var grid = $("feedbackCardsGrid");
    if (!grid) return;
    var starsStr = "★★★★★".slice(0, item.stars) + "☆☆☆☆☆".slice(0, 5 - item.stars);
    var initials = item.name.split(" ").map(function (w) { return w[0]; }).join("").slice(-2).toUpperCase() || "KH";

    var card = document.createElement("div");
    card.className = "feedback-card";
    card.innerHTML =
      '<div class="feedback-card-top">' +
        '<div class="feedback-stars">' + starsStr + '</div>' +
        '<span class="feedback-tag feedback-tag-new">Mới gửi 🌟</span>' +
      '</div>' +
      '<p class="feedback-quote">"' + escHtml(item.comment) + '"</p>' +
      '<div class="feedback-author-row">' +
        '<div class="feedback-avatar" style="background:#fee2e2;color:#be123c;">' + escHtml(initials) + '</div>' +
        '<div class="feedback-author-info">' +
          '<strong>' + escHtml(item.name) + '</strong>' +
          '<span>' + escHtml(item.company) + ' • <em style="color:#e11d48;">' + escHtml(item.category) + '</em></span>' +
        '</div>' +
      '</div>';
    grid.insertBefore(card, grid.firstChild);
  }

  function loadPersistedFeedbacks() {
    try {
      var list = JSON.parse(localStorage.getItem("namviet_feedbacks") || "[]");
      list.slice(0, 5).reverse().forEach(function (item) {
        appendFeedbackCardToGrid(item);
      });
    } catch (err) {}
  }

  document.addEventListener("DOMContentLoaded", init);
})();
