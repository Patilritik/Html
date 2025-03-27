console.log("Before Office.onReady");

Office.onReady(async () => {
    console.log("Office.onReady triggered");

    if (!localStorage.getItem("loginData")) {
        console.log("No login data found");
        window.location.href = "index.html";
        return;
    }

    const loginData = JSON.parse(localStorage.getItem("loginData"));
    const UserId = loginData.loginid;
    const ComCode = loginData.comcode;
    const status = loginData.status;
    const Token = loginData.Token;

    const departmentSelect = document.getElementById("department");
    const agreementTypeSelect = document.getElementById("agreementType");
    const proceedBtn = document.getElementById("proceedBtn");
    const loader = document.getElementById("loader");

    function showLoader() {
        loader.style.display = "block";
    }

    function hideLoader() {
        loader.style.display = "none";
    }

    async function fetchDepartmentList() {
        const apiUrl = `https://lapi.convergelego.com/api/AddLegalAgreement/Departmentlisit?companycode=${ComCode}&status=${status}`;
        try {
            const response = await fetch(apiUrl, {
                method: "GET",
                headers: {
                    "Userid": UserId,
                    "Key": "0",
                    "Comcode": ComCode
                }
            });

            const result = await response.json();
            departmentSelect.innerHTML = "";

            if (result.status === 200) {
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Department";
                departmentSelect.appendChild(defaultOption);

                const departments = result?.Detail?.data || [];
                departments.forEach(dept => {
                    const option = document.createElement("option");
                    option.value = dept.DeptId;
                    option.textContent = dept.DeptName;
                    departmentSelect.appendChild(option);
                });

                if (loginData.Depid) {
                    departmentSelect.value = loginData.Depid;
                    fetchAgreementTypeList(loginData.Depid);
                }
            } else {
                departmentSelect.innerHTML = `<option value="">No departments found</option>`;
            }
        } catch (err) {
            console.error("Department fetch error:", err);
            departmentSelect.innerHTML = `<option value="">No departments found</option>`;
        }
    }

    async function fetchAgreementTypeList(deptId) {
        const apiUrl = `https://lapi.convergelego.com/api/STDAgreementType/MstAgTypelisit?comcode=${ComCode}&depid=${deptId}`;
        try {
            agreementTypeSelect.disabled = false;
            const response = await fetch(apiUrl, {
                method: "GET",
                headers: {
                    "Userid": UserId,
                    "Key": "0",
                    "Comcode": ComCode
                }
            });

            const result = await response.json();
            agreementTypeSelect.innerHTML = "";

            if (result.status === 200) {
                const defaultOption = document.createElement("option");
                defaultOption.value = "";
                defaultOption.textContent = "Select Agreement Type";
                agreementTypeSelect.appendChild(defaultOption);

                const types = result?.Detail?.data || [];
                types.forEach(ag => {
                    const option = document.createElement("option");
                    option.value = ag.Agtypeid;
                    option.textContent = ag.AgtypeDesc;
                    agreementTypeSelect.appendChild(option);
                });
            } else {
                agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
            }
        } catch (err) {
            console.error("Agreement type fetch error:", err);
            agreementTypeSelect.innerHTML = `<option value="">No agreement types found</option>`;
        }

        updateProceedButtonState();
    }

    function updateProceedButtonState() {
        proceedBtn.disabled = !(departmentSelect.value && agreementTypeSelect.value);
    }

    departmentSelect.addEventListener("change", (e) => {
        const deptId = e.target.value;
        if (deptId) {
            fetchAgreementTypeList(deptId);
        } else {
            agreementTypeSelect.innerHTML = `<option value="">Select a department first</option>`;
            agreementTypeSelect.disabled = true;
            proceedBtn.disabled = true;
        }
    });

    agreementTypeSelect.addEventListener("change", updateProceedButtonState);

    proceedBtn.addEventListener("click", async () => {
        const deptId = departmentSelect.value;
        const agTypeId = agreementTypeSelect.value;
        const apiUrl = "https://addinapi.convergelego.com/api/CompanyMaster/GetMstCauseLisit";

        showLoader(); // Show loader on Proceed
        try {
            const response = await fetch(apiUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Userid": UserId,
                    "Key": Token,
                    "Token": Token,
                    "Comcode": ComCode
                },
                body: JSON.stringify({
                    deptid: deptId,
                    agrid: agTypeId,
                    statusid: "1"
                })
            });

            const result = await response.json();
            renderClausesTable(result?.Detail?.data || []);
        } catch (error) {
            console.error("Clause fetch error:", error);
            renderClausesTable([]);
        } finally {
            hideLoader(); // Hide loader after fetch
        }
    });

    function renderClausesTable(clauses) {
        const container = document.getElementById("clausesTableContainer");
        const tbody = document.getElementById("clausesTableBody");

        tbody.innerHTML = "";

        if (!clauses.length) {
            tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;">No clauses found.</td></tr>`;
            container.style.display = "block";
            return;
        }

        clauses.forEach(c => {
            const row = document.createElement("tr");
            row.innerHTML = `
                <td>${c.id || '-'}</td>
                <td>${c.causetitle || '-'}</td>
                <td style="white-space: pre-wrap; max-width: 300px;">${c.cause || '-'}</td>
                <td>${c.crby || '-'}</td>
                <td>${c.cron || '-'}</td>
                `;
                tbody.appendChild(row);
                // <td>${c.statusdesc || '-'}</td>
        });

        container.style.display = "block";
    }

    // Initial fetch
    await fetchDepartmentList();
});
