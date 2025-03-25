Office.onReady(() => {
    // Retrieve login data from localStorage
    const loginData = JSON.parse(localStorage.getItem("loginData"));
    console.log("Login dataaaaaaaaa:", loginData);
    const userInfoDiv = document.getElementById("userInfo");
    const departmentSelect = document.getElementById("department");
    const agreementTypeSelect = document.getElementById("agreementType");

    // Check if login data exists
    if (!loginData) {
        userInfoDiv.innerHTML = "<p>Please log in first.</p>";
        window.location.href = "index.html"; // Redirect to login if no data
        return;
    }

    // Display login data in the userInfo section
    userInfoDiv.innerHTML = `
        <p><strong>Login ID:</strong> ${loginData.loginid}</p>
        <p><strong>Employee Name:</strong> ${loginData.Empname}</p>
        <p><strong>Company Code:</strong> ${loginData.comcode}</p>
        <p><strong>Company Name:</strong> ${loginData.CompanyName}</p>
        <p><strong>Department ID:</strong> ${loginData.Depid}</p>
        <p><strong>User Type:</strong> ${loginData.Utype}</p>
    `;

    // Populate department dropdown with Depid from login data
    departmentSelect.innerHTML = ""; // Clear "Loading..." option
    departmentSelect.innerHTML = `<option value="${loginData.Depid}">${loginData.Depid}</option>`;
    // If you need a full list of departments, fetch them with an API call here

    // Populate agreementType dropdown (placeholder)
    agreementTypeSelect.innerHTML = ""; // Clear "Loading..." option
    agreementTypeSelect.innerHTML = `<option value="">Select Agreement Type</option>`;
    // Add API call here if you have agreement types to fetch
});