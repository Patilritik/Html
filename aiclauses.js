// aiclauses.js

Office.onReady(async () => {
    const loginData = JSON.parse(localStorage.getItem("loginData"));
    if (!loginData?.Token) {
      alert("Authentication error: Please log in again.");
      window.location.href = "index.html";
      return;
    }
  
    const apiToken = loginData.Token;
    const searchInput = document.getElementById("searchInput");
    const searchButton = document.getElementById("searchButton");
    const optimizeBtn = document.querySelector(".optimize-btn");
    const contentDiv = document.querySelector(".content-div p");
    const addToWordBtn = document.querySelector(".content-div button");
  
    let generatedText = "";
  
    searchButton.addEventListener("click", async () => {
      const query = searchInput.value.trim();
      if (!query) return alert("Please enter a query to search.");
  
      const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: "gpt-4o-mini",
          messages: [
            {
              role: "system",
              content:
                "You are an expert in real estate law. Respond strictly with relevant legal information and clauses related to real estate and property. Do not include unrelated information."
            },
            {
              role: "user",
              content: query
            }
          ],
          max_tokens: 1000,
          temperature: 0.5
        })
      });
  
      const result = await response.json();
      generatedText = result?.choices?.[0]?.message?.content || "No content generated.";
      contentDiv.textContent = generatedText;
    });
  
    optimizeBtn.addEventListener("click", async () => {
      if (!generatedText) return alert("Please generate text first.");
  
      const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: "gpt-4o-mini",
          messages: [
            {
              role: "system",
              content: "You are an expert at optimizing and improving text clarity."
            },
            {
              role: "user",
              content: generatedText
            }
          ],
          max_tokens: 1000,
          temperature: 0.5
        })
      });
  
      const result = await response.json();
      generatedText = result?.choices?.[0]?.message?.content || "Optimization failed.";
      contentDiv.textContent = generatedText;
    });
  
    addToWordBtn.addEventListener("click", async () => {
      const confirmInsert = confirm("Do you want to add this clause in Word?");
      if (!confirmInsert || !generatedText) return;
  
      try {
        await Word.run(async (context) => {
          const body = context.document.body;
          const para = body.insertParagraph(generatedText, Word.InsertLocation.end);
          para.font.size = 14;
          await context.sync();
        });
      } catch (error) {
        console.error("Error inserting into Word:", error);
      }
    });
  });
  