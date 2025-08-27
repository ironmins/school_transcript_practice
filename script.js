// script.js
document.getElementById("analyzeBtn").addEventListener("click", async () => {
  const fileInput = document.getElementById("fileInput");
  if (!fileInput.files.length) {
    alert("파일을 업로드해주세요.");
    return;
  }
  // 분석 로직 예시
  const analysisArea = document.getElementById("analysisArea");
  analysisArea.innerHTML = "<h2>과목별 분석</h2><div>분석 결과 예시...</div>";
  document.getElementById("saveResultBtn").style.display = "inline-block";
});

document.getElementById("saveResultBtn").addEventListener("click", () => {
  const analysisHTML = `
  <!DOCTYPE html>
  <html lang="ko">
  <head>
    <meta charset="UTF-8">
    <title>성적 분석 결과</title>
  </head>
  <body>
    ${document.getElementById("analysisArea").innerHTML}
  </body>
  </html>
  `;
  const blob = new Blob([analysisHTML], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "analysis_result.html";
  a.click();
  URL.revokeObjectURL(url);
});
