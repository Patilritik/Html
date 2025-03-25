console.log("Before Office.onReady");

Office.onReady(async () => {
    console.log("Office.onReady triggered");

    // Retrieve login data from localStorage
    if (!localStorage.getItem("loginData")) {
        console.log("No login data found");
        window.location.href = "index.html"; // Redirect to login if no data
        return;
    }

    const loginData = JSON.parse(localStorage.getItem("loginData"));
    const UserId = loginData.loginid;
    const ComCode = loginData.comcode;
    const status = loginData.status;
    console.log("Login data:", loginData);

    const userInfoDiv = document.getElementById("userInfo");
    const departmentSelect = document.getElementById("department");
    const agreementTypeSelect = document.getElementById("agreementType");

    // Check if login data exists
    if (!loginData) {
        userInfoDiv.innerHTML = "<p>Please log in first.</p>";
        window.location.href = "index.html"; // Redirect to login if no data
        return;
    }

    // Fetch department list from API
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
        console.log("Department list response:", result);

        // Clear "Loading..." option
        departmentSelect.innerHTML = "";

        // Check if result.status == 200
        if (result.status == 200) {
            // Add a default "Select Department" option
            const defaultOption = document.createElement("option");
            defaultOption.value = "";
            defaultOption.textContent = "Select Department";
            departmentSelect.appendChild(defaultOption);

            // Assuming the department list is in result.data (adjust based on actual API response structure)
            const departments = result.data || [];
            if (Array.isArray(departments) && departments.length > 0) {
                // Populate dropdown with department list
                departments.forEach(department => {
                    const option = document.createElement("option");
                    option.value = department.DeptId; // Use DeptId as the value
                    option.textContent = department.DeptName || department.DeptId; // Use DeptName if available, else DeptId
                    departmentSelect.appendChild(option);
                });

                // Optionally pre-select the user's department if available in loginData
                if (loginData.Depid) {
                    departmentSelect.value = loginData.Depid;
                }
            } else {
                departmentSelect.innerHTML = `<option value="">No departments found</option>`;
            }
        } else {
            departmentSelect.innerHTML = `<option value="">No departments found</option>`;
        }
    } catch (error) {
        console.error("Error fetching department list:", error);
        departmentSelect.innerHTML = `<option value="">No departments found</option>`;
    }

    // Populate agreementType dropdown (placeholder for now)
    agreementTypeSelect.innerHTML = ""; // Clear "Loading..." option
    agreementTypeSelect.innerHTML = `<option value="">Select Agreement Type</option>`;
    // Add API call for agreement types if needed
});