class ScoreAnalyzer {
  constructor() {
    this.originalData = [];
    this.selectedSubjects = [];
    this.subjectStats = {};
    this.gradeStats = {};
    this.studentStats = {};
  }

  initializeEventListeners() {
    const exportButton = document.getElementById('exportHtml');
    if (exportButton) {
      exportButton.addEventListener('click', () => this.exportAsHtml());
    }
  }

  exportAsHtml() {
    const newWindow = window.open('', '_blank');
    if (!newWindow) return;

    const serializedData = JSON.stringify({
      originalData: this.originalData,
      selectedSubjects: this.selectedSubjects,
      subjectStats: this.subjectStats,
      gradeStats: this.gradeStats,
      studentStats: this.studentStats
    });

    const htmlContent = \`
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>공유용 성적 분석</title>
  <link rel="stylesheet" href="https://ironmins.github.io/school_transcript_practice/style.css" />
</head>
<body>
  <div id="main-content">
    <h1>성적 분석 결과</h1>
    <div id="subjectTable"></div>
    <div id="gradeChart"></div>
    <div id="studentAnalysis"></div>
  </div>
  <script src="https://ironmins.github.io/school_transcript_practice/script.js"></script>
  <script>
    const sharedAnalysisData = \${serializedData};
    document.addEventListener('DOMContentLoaded', () => {
      const analyzer = new ScoreAnalyzer();
      analyzer.initializeFromSharedData(sharedAnalysisData);
    });
  </script>
</body>
</html>
\`;
    newWindow.document.write(htmlContent);
    newWindow.document.close();
  }

  initializeFromSharedData(serializedData) {
    const data = typeof serializedData === 'string' ? JSON.parse(serializedData) : serializedData;
    this.originalData = data.originalData;
    this.selectedSubjects = data.selectedSubjects;
    this.subjectStats = data.subjectStats;
    this.gradeStats = data.gradeStats;
    this.studentStats = data.studentStats;

    this.renderSubjectTable();
    this.renderGradeChart();
    this.renderStudentAnalysis();
  }

  renderSubjectTable() {
    const container = document.getElementById('subjectTable');
    if (!container) return;
    container.innerHTML = '<p>[과목별 분석 데이터 출력]</p>';
  }

  renderGradeChart() {
    const container = document.getElementById('gradeChart');
    if (!container) return;
    container.innerHTML = '<p>[평균등급 분포 차트 출력]</p>';
  }

  renderStudentAnalysis() {
    const container = document.getElementById('studentAnalysis');
    if (!container) return;
    container.innerHTML = '<p>[학생별 분석 내용 출력]</p>';
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const analyzer = new ScoreAnalyzer();
  analyzer.initializeEventListeners();
});