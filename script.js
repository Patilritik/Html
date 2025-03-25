Office.onReady(() => {
    const params = new URLSearchParams(window.location.search);
    const type = params.get("type");
  
    document.getElementById("loginForm").onsubmit = function (e) {
      e.preventDefault();
      // Simulated login
      const username = document.getElementById("username").value;
      const password = document.getElementById("password").value;
  
      if (username && password) {
        showContent(type);
      } else {
        alert("Please enter credentials.");
      }
    };
  });
  
  function showContent(type) {
    document.getElementById("loginForm").style.display = "none";
    const content = document.getElementById("content");
    const title = document.getElementById("contentTitle");
    const input = document.getElementById("clauseInput");
  
    content.style.display = "block";
    title.textContent = type === "ai" ? "AI Clauses Generator" : "Manual Clauses Entry";
    input.placeholder = type === "ai" ? "Generate or enter AI-based clause" : "Enter your clause manually";
  }
  
  function insertClause() {
    const text = document.getElementById("clauseInput").value;
    if (!text) return alert("Please enter a clause.");
  
    Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(text + "\n\n", Word.InsertLocation.end);
      await context.sync();
      document.getElementById("status").textContent = "Clause inserted successfully!";
    }).catch((err) => {
      console.error("Error inserting clause:", err);
    });
  }
  