class ScoreAnalyzer {
    constructor() {
        this.filesData = new Map(); // 파일명 -> 분석 데이터 매핑
        this.combinedData = null; // 통합된 분석 데이터
        this.selectedFiles = null; // 사용자가 선택/드롭한 파일 목록
        this.initializeEventListeners();
        // 🔹 URL 해시에 공유 데이터가 있으면 복원
        const hash = window.location.hash || '';
        const m = hash.match(/data=([^&]+)/);
        if (m && window.LZString) {
            try {
                const decoded = window.LZString.decompressFromEncodedURIComponent(m[1]);
                if (decoded) {
                    window.PRELOADED_DATA = JSON.parse(decoded);
                }
            } catch (e) {
                console.warn('해시 데이터 복원 실패:', e);
            }
        }


        // If the page provides preloaded analysis data, render directly
        if (window.PRELOADED_DATA) {
            try {
                this.combinedData = window.PRELOADED_DATA;
                const upload = document.querySelector('.upload-section');
                if (upload) upload.style.display = 'none';
                const results = document.getElementById('results');
                if (results) results.style.display = 'block';
                this.displayResults();
            // 내보내기 버튼 활성화
            const shareLinkBtn = document.getElementById('shareLinkBtn');
            const exportSingleBtn = document.getElementById('exportSingleBtn');
            const exportZipBtn = document.getElementById('exportZipBtn');
            if (shareLinkBtn) shareLinkBtn.disabled = false;
            if (exportSingleBtn) exportSingleBtn.disabled = false;
            if (exportZipBtn) exportZipBtn.disabled = false;
                const exportBtn = document.getElementById('exportBtn');
                if (exportBtn) exportBtn.disabled = false;
            } catch (e) {
                console.error('PRELOADED_DATA 처리 중 오류:', e);
            }
        }
    }

    initializeEventListeners() {
  const shareLinkBtn = document.getElementById('shareLinkBtn');
  const exportSingleBtn = document.getElementById('exportSingleBtn');
  const exportZipBtn = document.getElementById('exportZipBtn');
        const fileInput = document.getElementById('excelFiles');
        const analyzeBtn = document.getElementById('analyzeBtn');
        const exportBtn = document.getElementById('exportBtn');
        const tabBtns = document.querySelectorAll('.tab-btn');
        const studentSearch = document.getElementById('studentSearch');
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');
        const showStudentDetail = document.getElementById('showStudentDetail');
        const tableViewBtn = document.getElementById('tableViewBtn');
        const detailViewBtn = document.getElementById('detailViewBtn');
        const pdfClassBtn = document.getElementById('pdfClassBtn');
        const uploadSection = document.querySelector('.upload-section');
        const fileLabel = document.querySelector('.file-input-label');

        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length > 0) {
                this.selectedFiles = files;
                this.displayFileList(files);
                analyzeBtn.disabled = false;
                this.hideError();
            }
        });

        // 새 버튼 리스너
        if (shareLinkBtn) shareLinkBtn.addEventListener('click', () => this.openShareLink());
        if (exportSingleBtn) exportSingleBtn.addEventListener('click', () => this.exportAsSingleHtml());
        if (exportZipBtn) exportZipBtn.addEventListener('click', () => this.exportAsHtml(true));

        // Drag & drop 지원 (업로드 섹션 전체)
        if (uploadSection) {
            const prevent = (ev) => {
                ev.preventDefault();
                ev.stopPropagation();
            };
            const setDragState = (on) => {
                if (fileLabel) fileLabel.classList.toggle('dragover', on);
                uploadSection.classList.toggle('dragover', on);
            };

            // 전역 기본 동작 방지: 페이지로 파일이 열리는 것을 방지
            ['dragover', 'drop'].forEach(evt => {
                window.addEventListener(evt, (ev) => {
                    prevent(ev);
                });
            });

            ['dragenter', 'dragover'].forEach(evt => {
                uploadSection.addEventListener(evt, (ev) => {
                    prevent(ev);
                    setDragState(true);
                });
            });
            ['dragleave', 'dragend'].forEach(evt => {
                uploadSection.addEventListener(evt, (ev) => {
                    prevent(ev);
                    setDragState(false);
                });
            });
            uploadSection.addEventListener('drop', (ev) => {
                prevent(ev);
                setDragState(false);
                const dropped = Array.from(ev.dataTransfer?.files || []);
                const files = dropped.filter(f => /\.(xlsx|xls)$/i.test(f.name));
                if (files.length === 0) {
                    this.showError('XLS/XLSX 파일을 드래그하여 업로드하세요.');
                    return;
                }
                this.selectedFiles = files;
                this.displayFileList(files);
                analyzeBtn.disabled = false;
                this.hideError();
                try { if (fileInput) fileInput.files = ev.dataTransfer.files; } catch (_) {}
            });
        }

        analyzeBtn.addEventListener('click', () => {
            this.analyzeFiles();
        });

        

        tabBtns.forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.switchTab(e.target.dataset.tab);
            });
        });

        studentSearch.addEventListener('input', (e) => {
            this.filterStudents(e.target.value);
        });

        gradeSelect.addEventListener('change', () => {
            this.updateClassOptions();
            this.updateStudentOptions();
        });

        classSelect.addEventListener('change', () => {
            this.updateStudentOptions();
        });

        studentSelect.addEventListener('change', () => {
            showStudentDetail.disabled = !studentSelect.value;
        });
        if (studentNameSearch) {
            studentNameSearch.addEventListener('input', () => {
                this.updateStudentOptions();
            });
        }

        showStudentDetail.addEventListener('click', () => {
            this.showStudentDetail();
        });

        tableViewBtn.addEventListener('click', () => {
            this.switchView('table');
        });

        detailViewBtn.addEventListener('click', () => {
            this.switchView('detail');
        });

        if (pdfClassBtn) {
            pdfClassBtn.addEventListener('click', () => this.generateSelectedClassPDF());
        }
    }

    displayFileList(files) {
        const fileList = document.getElementById('fileList');
        fileList.innerHTML = '<h4>선택된 파일:</h4>';
        
        const ul = document.createElement('ul');
        files.forEach(file => {
            const li = document.createElement('li');
            li.textContent = file.name;
            ul.appendChild(li);
        });
        
        fileList.appendChild(ul);
        fileList.style.display = 'block';
    }

    async analyzeFiles() {
        const fileInput = document.getElementById('excelFiles');
        const files = (this.selectedFiles && this.selectedFiles.length > 0)
            ? this.selectedFiles
            : Array.from(fileInput.files);
        
        if (files.length === 0) {
            this.showError('파일을 선택해주세요.');
            return;
        }

        this.showLoading();
        
        try {
            this.filesData.clear();
            
            for (const file of files) {
                const data = await this.readExcelFile(file);
                const fileData = this.parseFileData(data, file.name);
                this.filesData.set(file.name, fileData);
            }
            
            this.combineAllData();
            this.displayResults();
            this.hideLoading();

            // Enable export button after successful analysis
            const exportBtn = document.getElementById('exportBtn');
            if (exportBtn) exportBtn.disabled = false;
        } catch (error) {
            this.hideLoading();
            this.showError('파일 분석 중 오류가 발생했습니다: ' + error.message);
        }
    }

    combineAllData() {
        if (this.filesData.size === 0) return;

        this.combinedData = {
            subjects: [],
            students: [],
            fileNames: Array.from(this.filesData.keys())
        };

        // 모든 과목을 통합 (중복 제거)
        const subjectMap = new Map();
        this.filesData.forEach((fileData) => {
            fileData.subjects.forEach(subject => {
                const key = `${subject.name}-${subject.credits}`;
                if (!subjectMap.has(key)) {
                    subjectMap.set(key, {
                        name: subject.name,
                        credits: subject.credits,
                        averages: [],
                        distributions: [],
                        columnIndex: subject.columnIndex
                    });
                }
                // 각 파일의 평균과 분포 저장
                subjectMap.get(key).averages.push(subject.average || 0);
                if (subject.distribution) {
                    subjectMap.get(key).distributions.push(subject.distribution);
                }
            });
        });

        // 과목별 전체 평균 계산
        subjectMap.forEach(subject => {
            subject.average = subject.averages.length > 0 
                ? subject.averages.reduce((sum, avg) => sum + avg, 0) / subject.averages.length 
                : 0;
            
            // 분포도 평균 계산
            if (subject.distributions.length > 0) {
                subject.distribution = {};
                const grades = ['A', 'B', 'C', 'D', 'E'];
                grades.forEach(grade => {
                    const values = subject.distributions
                        .map(dist => dist[grade] || 0)
                        .filter(val => val > 0);
                    subject.distribution[grade] = values.length > 0 
                        ? values.reduce((sum, val) => sum + val, 0) / values.length 
                        : 0;
                });
            }
        });

        this.combinedData.subjects = Array.from(subjectMap.values());

        // 모든 학생 데이터 통합
        let studentCounter = 1;
        this.filesData.forEach((fileData, fileName) => {
            fileData.students.forEach(student => {
                const fileNameParts = fileName.split('.')[0];
                
                const combinedStudent = {
                    ...student,
                    number: studentCounter++,
                    originalNumber: student.number,
                    originalName: student.name, // 원본 이름 보존
                    fileName: fileName,
                    name: student.name, // 실제 학생 이름 사용
                    displayName: `${fileData.grade}학년${fileData.class}반-${student.name}`, // 표시용 이름
                    grade: fileData.grade, // 파일의 A3 셀에서 추출한 학년
                    class: fileData.class, // 파일의 A3 셀에서 추출한 반
                    percentiles: {}
                };
                this.combinedData.students.push(combinedStudent);
            });
        });

        // 과목별 백분위 계산
        this.calculatePercentiles();
        
        // 평균등급 기준 순위 계산
        this.calculateAverageGradeRanks();
    }

    calculatePercentiles() {
        if (!this.combinedData) return;

        this.combinedData.subjects.forEach(subject => {
            // 해당 과목의 석차가 있는 모든 학생 수집
            const studentsWithRanks = this.combinedData.students
                .filter(student => {
                    const rank = student.ranks[subject.name];
                    return rank !== undefined && rank !== null && !isNaN(rank);
                })
                .map(student => ({
                    student: student,
                    rank: student.ranks[subject.name]
                }))
                .sort((a, b) => a.rank - b.rank); // 석차 순으로 정렬

            if (studentsWithRanks.length === 0) return;

            const totalStudents = studentsWithRanks.length;

            // 각 학생의 백분위 계산
            studentsWithRanks.forEach((item, index) => {
                const studentRank = item.rank;
                
                // 같은 석차의 학생들 찾기
                const sameRankStudents = studentsWithRanks.filter(s => s.rank === studentRank);
                const sameRankCount = sameRankStudents.length;
                
                // 해당 석차보다 나쁜 석차의 학생 수 (석차가 높은 학생들)
                const worseRankCount = studentsWithRanks.filter(s => s.rank > studentRank).length;
                
                // 백분위 계산: (더 나쁜 석차 학생 수 + 동점자의 절반) / 전체 학생 수 * 100
                // 이렇게 하면 1등(rank=1)이 가장 높은 백분위를 갖게 됨
                const percentile = ((worseRankCount + (sameRankCount - 1) / 2) / totalStudents) * 100;
                
                // 0~100 범위로 제한하고 반올림
                const finalPercentile = Math.max(0, Math.min(100, Math.round(percentile)));
                
                item.student.percentiles[subject.name] = finalPercentile;
            });
        });
    }

    calculateAverageGradeRanks() {
        if (!this.combinedData) return;

        // 평균등급이 있는 학생들만 필터링하고 정렬
        const studentsWithGrades = this.combinedData.students
            .filter(student => student.weightedAverageGrade !== null && student.weightedAverageGrade !== undefined)
            .sort((a, b) => a.weightedAverageGrade - b.weightedAverageGrade);

        if (studentsWithGrades.length === 0) return;

        let currentRank = 1;
        let previousGrade = null;
        let sameGradeCount = 0;

        studentsWithGrades.forEach((student, index) => {
            const studentGrade = student.weightedAverageGrade;
            
            // 이전 학생과 평균등급이 다르면 순위 업데이트
            if (previousGrade !== null && Math.abs(studentGrade - previousGrade) >= 0.01) {
                currentRank = index + 1;
                sameGradeCount = 1;
            } else if (previousGrade !== null) {
                // 같은 등급
                sameGradeCount++;
            } else {
                // 첫 번째 학생
                sameGradeCount = 1;
            }
            
            // 같은 평균등급의 학생 수 계산
            const totalSameGrade = studentsWithGrades.filter(s => 
                Math.abs(s.weightedAverageGrade - studentGrade) < 0.01
            ).length;
            
            student.averageGradeRank = currentRank;
            student.sameGradeCount = totalSameGrade;
            student.totalGradedStudents = studentsWithGrades.length;
            
            previousGrade = studentGrade;
        });

        // 평균등급이 없는 학생들은 순위도 null로 설정
        this.combinedData.students.forEach(student => {
            if (student.weightedAverageGrade === null || student.weightedAverageGrade === undefined) {
                student.averageGradeRank = null;
                student.sameGradeCount = null;
            }
            
            // 9등급 환산 평균 계산 (기존 데이터에 없는 경우)
            if (!student.weightedAverage9Grade) {
                student.weightedAverage9Grade = this.calculateWeightedAverage9Grade(student, this.combinedData.subjects);
            }
        });
    }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('파일 읽기 실패'));
            reader.readAsArrayBuffer(file);
        });
    }

    parseFileData(data, fileName) {
        const fileData = {
            fileName: fileName,
            data: data,
            subjects: [],
            students: [],
            grade: 1,
            class: 1
        };

        // A3 셀에서 학년/반 정보 추출 (0-based index로는 행 2, 열 0)
        if (data[2] && data[2][0]) {
            const classInfo = data[2][0].toString();
            console.log('A3 셀 내용:', classInfo); // 디버깅용
            
            // "학년도" 뒤에 오는 학년 정보와 "반" 앞에 오는 반 정보 추출
            // 예: "2025학년도   1학기   주간      1학년     4반"
            const gradeMatch = classInfo.match(/\s+(\d+)학년/);
            const classMatch = classInfo.match(/\s+(\d+)반/);
            
            if (gradeMatch) {
                fileData.grade = parseInt(gradeMatch[1]);
                console.log('추출된 학년:', fileData.grade); // 디버깅용
            }
            if (classMatch) {
                fileData.class = parseInt(classMatch[1]);
                console.log('추출된 반:', fileData.class); // 디버깅용
            }
        }

        // 과목명 추출 (행 4, D열부터) - 0-based index로는 행 3
        const subjectRow = data[3]; // 행 4
        for (let i = 3; i < subjectRow.length; i++) { // D열부터
            const cellValue = subjectRow[i];
            if (cellValue && typeof cellValue === 'string' && cellValue.includes('(')) {
                const match = cellValue.match(/^(.+)\((\d+)\)$/);
                if (match) {
                    fileData.subjects.push({
                        name: match[1].trim(),
                        credits: parseInt(match[2]),
                        columnIndex: i,
                        scores: []
                    });
                }
            }
        }

        // 과목별 평균 (행 5) - 0-based index로는 행 4
        const averageRow = data[4];
        fileData.subjects.forEach(subject => {
            const avgValue = averageRow[subject.columnIndex];
            subject.average = avgValue ? parseFloat(avgValue) : 0;
        });

        // 성취도 분포 (행 6) - 0-based index로는 행 5
        const distributionRow = data[5];
        this.parseAchievementDistribution(distributionRow, fileData.subjects);

        // 학생 데이터 파싱 (행 7부터 시작, 5행씩 묶여있음)
        this.parseStudentData(data, fileData);

        return fileData;
    }

    parseAchievementDistribution(distributionRow, subjects) {
        subjects.forEach(subject => {
            subject.distribution = {};
            const cellValue = distributionRow[subject.columnIndex];
            
            if (cellValue && typeof cellValue === 'string') {
                // "A(6.3)B(15.3)C(12.6)D(18.9)E(46.8)" 형식에서 각 등급과 비율 추출
                const gradeMatches = cellValue.match(/[ABCDE]\(\d+\.?\d*\)/g);
                if (gradeMatches) {
                    gradeMatches.forEach(match => {
                        const gradeMatch = match.match(/([ABCDE])\((\d+\.?\d*)\)/);
                        if (gradeMatch) {
                            const grade = gradeMatch[1];
                            const percentage = parseFloat(gradeMatch[2]);
                            subject.distribution[grade] = percentage;
                        }
                    });
                }
            }
        });
    }

    parseStudentData(data, fileData) {
        // 학생 데이터는 행 7부터 시작해서 각 학생마다 5행씩 사용
        // 행 7: 번호 + 합계(원점수)
        // 행 8: 성취도
        // 행 9: 석차등급  
        // 행 10: 석차
        // 행 11: 수강자수
        
        let consecutiveEmptyRows = 0;
        const maxConsecutiveEmpty = 15; // 연속으로 15행이 비어있으면 종료
        
        for (let i = 6; i < data.length; i += 5) { // 0-based로 행 7부터, 5행씩 건너뛰기
            const scoreRow = data[i];     // 합계(원점수) 행
            const achievementRow = data[i + 1]; // 성취도 행
            const gradeRow = data[i + 2];       // 석차등급 행
            const rankRow = data[i + 3];        // 석차 행
            const totalRow = data[i + 4];       // 수강자수 행
            
            // 학생 번호가 있는지 확인 (A열)
            if (!scoreRow || !scoreRow[0] || isNaN(scoreRow[0])) {
                consecutiveEmptyRows += 5; // 5행씩 건너뛰므로 5 증가
                if (consecutiveEmptyRows >= maxConsecutiveEmpty) {
                    console.log(`연속으로 ${consecutiveEmptyRows}행이 비어있어 파싱을 종료합니다. (행 ${i + 1})`);
                    break;
                }
                continue; // 빈 행은 건너뛰고 다음 학생 찾기
            }
            
            // 유효한 학생 데이터를 찾았으면 연속 빈 행 카운터 리셋
            consecutiveEmptyRows = 0;
            
            console.log(`학생 발견: 행 ${i + 1}, 번호: ${scoreRow[0]}, 이름: ${scoreRow[1] || '미기입'}`);
            
            const student = {
                number: scoreRow[0],
                name: scoreRow[1] || `학생${scoreRow[0]}`, // B열에서 학생 이름 추출
                scores: {},
                achievements: {},
                grades: {},
                ranks: {},
                totalStudents: null
            };

            // 각 과목별 데이터 추출
            fileData.subjects.forEach(subject => {
                const colIndex = subject.columnIndex;
                
                // 점수 (원점수 추출)
                if (scoreRow[colIndex]) {
                    const scoreText = scoreRow[colIndex].toString();
                    const scoreMatch = scoreText.match(/(\d+\.?\d*)\((\d+)\)/);
                    if (scoreMatch) {
                        student.scores[subject.name] = parseFloat(scoreMatch[2]); // 원점수
                    }
                }
                
                // 성취도
                if (achievementRow && achievementRow[colIndex]) {
                    student.achievements[subject.name] = achievementRow[colIndex];
                }
                
                // 석차등급
                if (gradeRow && gradeRow[colIndex] && !isNaN(gradeRow[colIndex])) {
                    student.grades[subject.name] = parseInt(gradeRow[colIndex]);
                }
                
                // 석차
                if (rankRow && rankRow[colIndex] && !isNaN(rankRow[colIndex])) {
                    student.ranks[subject.name] = parseInt(rankRow[colIndex]);
                }
                
                // 수강자수 (첫 번째 과목에서만 가져오기)
                if (!student.totalStudents && totalRow && totalRow[colIndex] && !isNaN(totalRow[colIndex])) {
                    student.totalStudents = parseInt(totalRow[colIndex]);
                }
            });

            // 가중평균등급 계산
            student.weightedAverageGrade = this.calculateWeightedAverageGrade(student, fileData.subjects);
            
            // 9등급 환산 평균 계산
            student.weightedAverage9Grade = this.calculateWeightedAverage9Grade(student, fileData.subjects);
            
            fileData.students.push(student);
        }
        
        console.log(`총 ${fileData.students.length}명의 학생 데이터를 파싱했습니다.`);
    }

    calculateWeightedAverageGrade(student, subjects) {
        let totalGradePoints = 0;
        let totalCredits = 0;
        
        subjects.forEach(subject => {
            const grade = student.grades[subject.name];
            if (grade && !isNaN(grade)) {
                totalGradePoints += grade * subject.credits;
                totalCredits += subject.credits;
            }
        });
        
        return totalCredits > 0 ? totalGradePoints / totalCredits : null;
    }

    calculateWeightedAveragePercentile(student, subjects) {
        let totalPercentilePoints = 0;
        let totalCredits = 0;
        
        // percentiles와 ranks 객체가 존재하는지 확인
        if (!student.percentiles || !student.ranks) {
            return null;
        }
        
        subjects.forEach(subject => {
            const percentile = student.percentiles[subject.name];
            const rank = student.ranks[subject.name];
            // 석차가 있는 과목만 계산에 포함 (석차 기준으로 백분위 계산했으므로)
            if (percentile !== undefined && percentile !== null && rank !== undefined && rank !== null && !isNaN(rank)) {
                totalPercentilePoints += percentile * subject.credits;
                totalCredits += subject.credits;
            }
        });
        
        return totalCredits > 0 ? totalPercentilePoints / totalCredits : null;
    }

    // 백분위를 9등급으로 환산하는 함수
    convertPercentileTo9Grade(percentile) {
        if (percentile === null || percentile === undefined || isNaN(percentile)) {
            return null;
        }
        
        if (percentile >= 96) return 1;  // 상위 4%
        if (percentile >= 89) return 2;  // 상위 11%
        if (percentile >= 77) return 3;  // 상위 23%
        if (percentile >= 60) return 4;  // 상위 40%
        if (percentile >= 40) return 5;  // 상위 60%
        if (percentile >= 23) return 6;  // 상위 77%
        if (percentile >= 11) return 7;  // 상위 89%
        if (percentile >= 4) return 8;   // 상위 96%
        return 9;                        // 하위 4%
    }

    // 9등급 가중평균 계산
    calculateWeightedAverage9Grade(student, subjects) {
        let totalGradePoints = 0;
        let totalCredits = 0;
        
        // percentiles와 ranks 객체가 존재하는지 확인
        if (!student.percentiles || !student.ranks) {
            return null;
        }
        
        subjects.forEach(subject => {
            const percentile = student.percentiles[subject.name];
            const rank = student.ranks[subject.name];
            // 석차가 있는 과목만 계산에 포함
            if (percentile !== undefined && percentile !== null && rank !== undefined && rank !== null && !isNaN(rank)) {
                const grade9 = this.convertPercentileTo9Grade(percentile);
                if (grade9 !== null) {
                    totalGradePoints += grade9 * subject.credits;
                    totalCredits += subject.credits;
                }
            }
        });
        
        return totalCredits > 0 ? totalGradePoints / totalCredits : null;
    }


    displayResults() {
        document.getElementById('results').style.display = 'block';
        this.displaySubjectAverages();
        this.displayGradeAnalysis();
        this.displayStudentAnalysis();
    }

    // Export a complete deployment package with all files
    async exportAsHtml(createFolder = true) {
        if (!this.combinedData) {
            this.showError('먼저 파일을 분석하세요.');
            return;
        }

        const timestamp = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        const folderName = `analysis_${timestamp.getFullYear()}${pad(timestamp.getMonth()+1)}${pad(timestamp.getDate())}_${pad(timestamp.getHours())}${pad(timestamp.getMinutes())}`;

        // Serialize current analysis data
        const dataJson = JSON.stringify(this.combinedData);

        // Helper to fetch text
        const safeFetchText = async (url) => {
            try {
                const res = await fetch(url, { cache: 'no-cache' });
                if (!res.ok) throw new Error('HTTP ' + res.status);
                return await res.text();
            } catch (e) {
                console.warn('리소스 로드 실패:', url, e);
                return '';
            }
        };

        // Get CSS content
        let cssContent = await safeFetchText('style.css');
        
        // CSS 내용 확인 및 디버깅
        console.log('CSS 내용 길이:', cssContent.length);
        if (!cssContent || cssContent.length < 100) {
            console.warn('CSS를 가져오지 못함, 대체 방법 사용');
            // style 태그에서 CSS 추출 시도
            const styleElement = document.querySelector('link[href="style.css"]');
            if (styleElement) {
                try {
                    const response = await fetch(styleElement.href);
                    cssContent = await response.text();
                } catch (e) {
                    console.error('CSS 대체 로드 실패:', e);
                    // 마지막 fallback - 기본 스타일 제공
                    cssContent = this.getFallbackCSS();
                }
            } else {
                cssContent = this.getFallbackCSS();
            }
        }

        // Get JS content and modify for standalone use
        let jsContent = await safeFetchText('script.js');
        console.log('JS 내용 길이:', jsContent.length);
        if (jsContent) {
            jsContent = this.createStandaloneScript(jsContent);
            console.log('수정된 JS 내용 길이:', jsContent.length);
        } else {
            console.error('JavaScript 파일을 로드할 수 없습니다');
            jsContent = this.getFallbackJS();
        }

        // Create HTML file content
        const htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>배포용 성적 분석 뷰어</title>
    <style>
        /* 메인 CSS */
        ${cssContent}
        
        /* 차트 대체 스타일 */
        .chart-placeholder {
            width: 100%;
            height: 350px;
            background: #f8f9fa;
            border: 2px dashed #dee2e6;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #6c757d;
            font-size: 1.1rem;
            border-radius: 8px;
            flex-direction: column;
            padding: 20px;
        }
        .chart-placeholder h4 {
            margin-bottom: 15px;
            color: #333;
        }
        .chart-placeholder p {
            margin: 5px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>성적 분석 결과 (배포용)</h1>
            <p>업로드 없이 저장된 분석 결과를 표시합니다</p>
        </header>
        <div class="upload-section" style="display:none;"></div>
        ${document.getElementById('results') ? document.getElementById('results').outerHTML : '<div id="results" class="results-section"></div>'}
        <div id="loading" class="loading" style="display:none;"></div>
        <div id="error" class="error-message" style="display:none;"></div>
        <footer class="app-footer">
            <div class="footer-right">
                <div class="credits">Made by NAMGUNG YEON (Seolak high school)</div>
                <a class="help-btn" href="https://namgungyeon.tistory.com/133" target="_blank" rel="noopener" title="도움말 보기">❔ 도움말</a>
            </div>
        </footer>
    </div>

    <script>
        // Preloaded analysis data embedded for offline viewing
        window.PRELOADED_DATA = ${dataJson};
    </script>
    <script src="script.js"></script>
</body>
</html>`;

        // Create ZIP file with JSZip (if available) or download files separately
        if (typeof JSZip !== 'undefined' && cssContent.length > 100) {
            // Use JSZip if available and CSS loaded successfully
            const zip = new JSZip();
            zip.file("index.html", htmlContent);
            zip.file("style.css", cssContent || "/* CSS 로드 실패 */");
            zip.file("script.js", jsContent || "/* JS 로드 실패 */");
            zip.file("README.txt", 
                "배포용 성적 분석 뷰어\\n" +
                "========================\\n\\n" +
                "사용법:\\n" +
                "1. index.html 파일을 웹브라우저에서 열어주세요\\n" +
                "2. 업로드 없이 바로 분석 결과를 확인할 수 있습니다\\n" +
                "3. index.html에 CSS가 내장되어 있어 단독으로 실행 가능합니다\\n\\n" +
                "파일 구성:\\n" +
                "- index.html: 메인 페이지 (CSS 내장)\\n" +
                "- style.css: 별도 스타일 파일 (참고용)\\n" +
                "- script.js: 분석 스크립트\\n\\n" +
                "Made by NAMGUNG YEON (Seolak high school)\\n" +
                "링크: https://namgungyeon.tistory.com/133"
            );
            
            const content = await zip.generateAsync({type: "blob"});
            const url = URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            a.download = folderName + ".zip";
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 0);
        } else {
            // Fallback: download files separately
            this.downloadFile(htmlContent, "index.html", "text/html");
            setTimeout(() => this.downloadFile(cssContent, "style.css", "text/css"), 500);
            setTimeout(() => this.downloadFile(jsContent, "script.js", "application/javascript"), 1000);
            setTimeout(() => {
                const readme = "배포용 성적 분석 뷰어\\n========================\\n\\n사용법:\\n1. 모든 파일을 같은 폴더에 저장하세요\\n2. index.html 파일을 웹브라우저에서 열어주세요\\n\\nMade by NAMGUNG YEON (Seolak high school)\\n링크: https://namgungyeon.tistory.com/133";
                this.downloadFile(readme, "README.txt", "text/plain");
            }, 1500);
            
            alert(`배포용 파일들을 다운로드하고 있습니다.\\n\\n모든 파일을 같은 폴더에 저장한 후\\nindex.html 파일을 열어서 사용하세요.`);
        }
    }
    downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType + ';charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 0);
    }

    getFallbackCSS() {
        // CSS 로드가 실패했을 때 사용할 기본 스타일
        return `
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: linear-gradient(180deg, #f7f9fc 0%, #eef2f7 100%);
    min-height: 100vh;
    padding: 20px;
}

.container {
    max-width: 1200px;
    margin: 0 auto;
    background: white;
    border-radius: 15px;
    box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
    overflow: hidden;
}

header {
    background: #8fbaf7;
    color: white;
    padding: 40px;
    text-align: center;
}

header h1 {
    font-size: 2.5rem;
    margin-bottom: 10px;
    font-weight: 300;
}

.results-section {
    padding: 40px;
}

.tabs {
    display: flex;
    border-bottom: 2px solid #eee;
    margin-bottom: 30px;
}

.tab-btn {
    flex: 1;
    padding: 15px 20px;
    background: none;
    border: none;
    cursor: pointer;
    font-size: 1rem;
    color: #666;
    transition: all 0.3s ease;
    border-bottom: 3px solid transparent;
}

.tab-btn.active {
    color: #4facfe;
    border-bottom-color: #4facfe;
    background: rgba(79, 172, 254, 0.05);
}

.tab-content {
    display: none;
}

.tab-content.active {
    display: block;
}

.tab-content h2 {
    color: #333;
    margin-bottom: 25px;
    font-size: 1.8rem;
    font-weight: 400;
}

.subject-averages {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 20px;
}

.subject-item {
    background: white;
    border-radius: 10px;
    padding: 25px;
    border-left: 5px solid #4facfe;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.08);
}

.students-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 20px;
}

.student-card {
    background: white;
    border-radius: 15px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    border: 1px solid rgba(0, 0, 0, 0.05);
    overflow: hidden;
}

.grade-analysis-container {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 30px;
}

.chart-section {
    background: #f8f9fa;
    border-radius: 10px;
    padding: 25px;
    text-align: center;
}

.stats-section {
    grid-column: 1 / -1;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 20px;
    background: #f8f9fa;
    border-radius: 10px;
    padding: 25px;
}

.stat-item {
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    background: white;
    border-radius: 8px;
    padding: 20px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

/* 하단 크레딧 푸터 (fallback) */
.app-footer {
    padding: 12px 40px 24px 40px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
}
.app-footer .footer-right {
    display: flex;
    align-items: center;
    gap: 10px;
}
.app-footer .credits {
    text-align: right;
    font-size: 0.85rem;
    color: #ffffff; /* 흰색으로 변경 */
    opacity: 0.95;
}
.app-footer .credits a:not(.help-btn) {
    color: #adb5bd;
    text-decoration: none;
    border-bottom: 1px dashed rgba(173,181,189,0.5);
}
.app-footer .credits a:not(.help-btn):hover {
    color: #6c757d;
    border-bottom-color: rgba(108,117,125,0.7);
}

/* 도움말 버튼 */
.help-btn {
    display: inline-block;
    padding: 6px 12px;
    font-size: 0.85rem;
    line-height: 1;
    border-radius: 999px;
    color: #4facfe;
    background: #ffffff;
    border: 1px solid #4facfe;
    text-decoration: none;
    transition: all 0.2s ease;
}
.help-btn:hover {
    color: #ffffff;
    background: #4facfe;
    box-shadow: 0 6px 16px rgba(79, 172, 254, 0.25);
}
`;
    }

    getFallbackJS() {
        // JavaScript 로드가 실패했을 때 사용할 기본 스크립트
        return `
class ScoreAnalyzer {
    constructor() {
        this.combinedData = window.PRELOADED_DATA || null;
        this.initializeEventListeners();
        
        if (this.combinedData) {
            console.log('사전 로드된 데이터 발견:', this.combinedData);
            const upload = document.querySelector('.upload-section');
            if (upload) upload.style.display = 'none';
            const results = document.getElementById('results');
            if (results) results.style.display = 'block';
            this.displayResults();
        }
    }
    
    initializeEventListeners() {
        const tabBtns = document.querySelectorAll('.tab-btn');
        tabBtns.forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.switchTab(e.target.dataset.tab);
            });
        });
    }
    
    switchTab(tabName) {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector('[data-tab="' + tabName + '"]').classList.add('active');

        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(tabName + '-tab').classList.add('active');
    }
    
    displayResults() {
        if (!this.combinedData) return;
        
        document.getElementById('results').style.display = 'block';
        this.displaySubjectAverages();
        this.displayGradeAnalysis();
        this.displayStudentAnalysis();
    }
    
    displaySubjectAverages() {
        const container = document.getElementById('subjectAverages');
        if (!container || !this.combinedData) return;
        
        container.innerHTML = '';
        this.combinedData.subjects.forEach(subject => {
            const div = document.createElement('div');
            div.className = 'subject-item';
            div.innerHTML = '<h3>' + subject.name + '</h3><p>평균: ' + (subject.average || 0).toFixed(1) + '점</p>';
            container.appendChild(div);
        });
    }
    
    displayGradeAnalysis() {
        // 간단한 통계만 표시
        const overallAvg = document.getElementById('overallAverage');
        const stdDev = document.getElementById('standardDeviation');
        
        if (this.combinedData && this.combinedData.students) {
            const grades = this.combinedData.students
                .filter(s => s.weightedAverageGrade)
                .map(s => s.weightedAverageGrade);
                
            if (grades.length > 0) {
                const avg = grades.reduce((sum, g) => sum + g, 0) / grades.length;
                if (overallAvg) overallAvg.textContent = avg.toFixed(2);
                
                const variance = grades.reduce((sum, g) => sum + Math.pow(g - avg, 2), 0) / grades.length;
                if (stdDev) stdDev.textContent = Math.sqrt(variance).toFixed(2);
            }
        }
        
        // 차트 대신 메시지 표시
        const scatterChart = document.getElementById('scatterChart');
        const barChart = document.getElementById('barChart');
        
        if (scatterChart && scatterChart.parentElement) {
            scatterChart.parentElement.innerHTML = '<div class="chart-placeholder"><h4>차트는 배포용에서 제외됨</h4><p>통계 정보는 위에서 확인하세요</p></div>';
        }
        
        if (barChart && barChart.parentElement) {
            barChart.parentElement.innerHTML = '<div class="chart-placeholder"><h4>차트는 배포용에서 제외됨</h4><p>통계 정보는 위에서 확인하세요</p></div>';
        }
    }
    
    displayStudentAnalysis() {
        // 기본적인 학생 목록만 표시
        const container = document.getElementById('studentTable');
        if (!container || !this.combinedData) return;
        
        container.innerHTML = '<p>학생 분석 데이터가 로드되었습니다. 총 ' + this.combinedData.students.length + '명</p>';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new ScoreAnalyzer();
});
`;
    }

    createStandaloneScript(originalScript) {
        // Chart.js 의존성을 제거하고 더 안전한 방식으로 변경
        let modifiedScript = originalScript;
        
        try {
            // 1. Chart.js 관련 전역 참조 제거
            modifiedScript = modifiedScript.replace(/Chart\.register\(.*?\);?/g, '// Chart.js 제거됨');
            modifiedScript = modifiedScript.replace(/ChartDataLabels/g, '{}');
            
            // 2. 차트 생성 메서드들을 간단한 플레이스홀더로 교체
            modifiedScript = modifiedScript.replace(
                /createScatterChart\([^{]*\{[^}]*\{[\s\S]*?\}\s*\}\s*\}/g,
                `createScatterChart(students) {
                    const ctx = document.getElementById('scatterChart');
                    if (!ctx || !ctx.parentElement) return;
                    ctx.parentElement.innerHTML = '<div class="chart-placeholder"><h4>산점도 차트</h4><p>배포용에서는 차트가 제외되었습니다</p></div>';
                }`
            );
            
            modifiedScript = modifiedScript.replace(
                /createGradeDistributionChart\([^{]*\{[^}]*\{[\s\S]*?\}\s*\}\s*\}/g,
                `createGradeDistributionChart(students) {
                    const ctx = document.getElementById('barChart');
                    if (!ctx || !ctx.parentElement) return;
                    ctx.parentElement.innerHTML = '<div class="chart-placeholder"><h4>분포 차트</h4><p>배포용에서는 차트가 제외되었습니다</p></div>';
                }`
            );
            
            modifiedScript = modifiedScript.replace(
                /createStudentPercentileChart\([^{]*\{[^}]*\{[\s\S]*?\}\s*\}\s*\}/g,
                `createStudentPercentileChart(student) {
                    const ctx = document.getElementById('studentPercentileChart');
                    if (!ctx || !ctx.parentElement) return;
                    ctx.parentElement.innerHTML = '<div class="chart-placeholder"><h4>학생별 차트</h4><p>배포용에서는 차트가 제외되었습니다</p></div>';
                }`
            );
            
            // 3. 차트 파괴 관련 코드 제거
            modifiedScript = modifiedScript.replace(/if \(this\.\w*Chart\) \{\s*this\.\w*Chart\.destroy\(\);\s*\}/g, '// 차트 파괴 코드 제거됨');
            
            // 4. new Chart 생성자 호출 제거
            modifiedScript = modifiedScript.replace(/this\.\w*Chart = new Chart\([^;]*\);/g, '// Chart 생성 제거됨');
            
            console.log('Chart.js 의존성 제거 완료');
            
        } catch (e) {
            console.error('스크립트 수정 중 오류 발생:', e);
            console.warn('기본 fallback 스크립트 사용');
            return this.getFallbackJS();
        }
        
        return modifiedScript;
    }

    displaySubjectAverages() {
        const container = document.getElementById('subjectAverages');
        container.innerHTML = '';

        if (!this.combinedData) return;

        this.combinedData.subjects.forEach(subject => {
            const subjectDiv = document.createElement('div');
            subjectDiv.className = 'subject-item';
            
            // 성취도 분포 HTML 생성
            let distributionHTML = '';
            if (subject.distribution) {
                distributionHTML = '<div class="achievement-bars">';
                Object.entries(subject.distribution).forEach(([grade, percentage]) => {
                    distributionHTML += `
                        <div class="achievement-bar">
                            <span class="achievement-label">${grade}</span>
                            <div class="achievement-bar-container">
                                <div class="achievement-bar-fill" style="width: ${percentage}%"></div>
                            </div>
                            <span class="achievement-percentage">${percentage.toFixed(1)}%</span>
                        </div>
                    `;
                });
                distributionHTML += '</div>';
            }
            
            subjectDiv.innerHTML = `
                <div class="subject-header">
                    <h3>${subject.name}</h3>
                    <span class="credits">${subject.credits}학점</span>
                </div>
                <div class="average-score">
                    <span class="score">${subject.average?.toFixed(1) || 'N/A'}</span>
                    <span class="label">평균 점수</span>
                </div>
                ${distributionHTML}
            `;
            container.appendChild(subjectDiv);
        });
    }


    displayGradeAnalysis() {
        if (!this.combinedData) return;

        // 평균등급이 있는 학생들만 필터링
        const studentsWithGrades = this.combinedData.students.filter(student => 
            student.weightedAverageGrade !== null
        );

        if (studentsWithGrades.length === 0) {
            return;
        }

        // 통계 계산
        const grades = studentsWithGrades.map(student => student.weightedAverageGrade);
        const overallAverage = grades.reduce((sum, grade) => sum + grade, 0) / grades.length;
        const variance = grades.reduce((sum, grade) => sum + Math.pow(grade - overallAverage, 2), 0) / grades.length;
        const standardDeviation = Math.sqrt(variance);
        const bestGrade = Math.min(...grades);
        const worstGrade = Math.max(...grades);

        // 통계 표시
        document.getElementById('overallAverage').textContent = overallAverage.toFixed(2);
        document.getElementById('standardDeviation').textContent = standardDeviation.toFixed(2);
        document.getElementById('bestGrade').textContent = bestGrade.toFixed(2);
        document.getElementById('worstGrade').textContent = worstGrade.toFixed(2);

        // 산점도 생성
        this.createScatterChart(studentsWithGrades);

        // 막대그래프 생성
        this.createGradeDistributionChart(studentsWithGrades);
    }

    createScatterChart(students) {
        const ctx = document.getElementById('scatterChart').getContext('2d');
        
        // 기존 차트가 있다면 파괴
        if (this.scatterChart) {
            this.scatterChart.destroy();
        }

        // 평균등급별로 학생을 정렬 (1등급부터 5등급 순)
        const sortedStudents = [...students].sort((a, b) => a.weightedAverageGrade - b.weightedAverageGrade);
        
        // 각 평균등급별로 같은 등급의 학생 수만큼 Y축에 분산
        const gradeGroups = {};
        students.forEach(student => {
            const grade = student.weightedAverageGrade.toFixed(2);
            if (!gradeGroups[grade]) {
                gradeGroups[grade] = [];
            }
            gradeGroups[grade].push(student);
        });

        const scatterData = [];
        Object.keys(gradeGroups).forEach(grade => {
            const studentsInGrade = gradeGroups[grade];
            studentsInGrade.forEach((student, index) => {
                // 같은 등급의 학생들을 Y축에서 약간씩 분산 (중앙 기준으로 ±0.05 범위)
                const yOffset = studentsInGrade.length > 1 
                    ? (index - (studentsInGrade.length - 1) / 2) * 0.02 
                    : 0;
                
                scatterData.push({
                    x: parseFloat(grade),
                    y: 0.5 + yOffset, // Y축 중앙(0.5) 기준으로 약간 분산
                    student: student
                });
            });
        });

        // 누적 비율 계산을 위한 데이터 생성
        const cumulativeData = [];
        const totalStudents = sortedStudents.length;
        
        // 0.1 단위로 등급 구간을 나누어 누적 비율 계산
        for (let grade = 1.0; grade <= 5.0; grade += 0.1) {
            const studentsUpToGrade = sortedStudents.filter(s => s.weightedAverageGrade <= grade).length;
            const cumulativePercentage = (studentsUpToGrade / totalStudents) * 100;
            
            cumulativeData.push({
                x: parseFloat(grade.toFixed(1)),
                y: cumulativePercentage
            });
        }

        this.scatterChart = new Chart(ctx, {
            type: 'scatter',
            data: {
                datasets: [{
                    label: '누적 비율',
                    type: 'line',
                    data: cumulativeData,
                    borderColor: 'rgba(231, 76, 60, 1)',
                    backgroundColor: 'rgba(231, 76, 60, 0.1)',
                    borderWidth: 3,
                    pointBackgroundColor: 'rgba(231, 76, 60, 1)',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    pointRadius: 5,
                    pointHoverRadius: 8,
                    fill: false,
                    tension: 0.3,
                    yAxisID: 'y1',
                    order: 1,
                    // 차트 영역 경계에서 점/선이 잘리지 않도록 여유를 둠
                    clip: 8
                }, {
                    label: '학생별 평균등급',
                    type: 'scatter',
                    data: scatterData,
                    backgroundColor: function(context) {
                        const grade = context.parsed.x;
                        if (grade <= 1.5) return 'rgba(26, 188, 156, 0.6)';
                        if (grade <= 2.0) return 'rgba(52, 152, 219, 0.6)';
                        if (grade <= 2.5) return 'rgba(155, 89, 182, 0.6)';
                        if (grade <= 3.0) return 'rgba(241, 196, 15, 0.6)';
                        if (grade <= 3.5) return 'rgba(230, 126, 34, 0.6)';
                        if (grade <= 4.0) return 'rgba(231, 76, 60, 0.6)';
                        if (grade <= 4.5) return 'rgba(189, 195, 199, 0.6)';
                        return 'rgba(127, 140, 141, 0.6)';
                    },
                    borderColor: function(context) {
                        const grade = context.parsed.x;
                        if (grade <= 1.5) return 'rgba(26, 188, 156, 0.8)';
                        if (grade <= 2.0) return 'rgba(52, 152, 219, 0.8)';
                        if (grade <= 2.5) return 'rgba(155, 89, 182, 0.8)';
                        if (grade <= 3.0) return 'rgba(241, 196, 15, 0.8)';
                        if (grade <= 3.5) return 'rgba(230, 126, 34, 0.8)';
                        if (grade <= 4.0) return 'rgba(231, 76, 60, 0.8)';
                        if (grade <= 4.5) return 'rgba(189, 195, 199, 0.8)';
                        return 'rgba(127, 140, 141, 0.8)';
                    },
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    borderWidth: 2,
                    pointHoverBorderWidth: 3,
                    yAxisID: 'y',
                    order: 2,
                    // 차트 영역 경계에서 점이 잘리지 않도록 여유를 둠
                    clip: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        top: 20,
                        bottom: 20,
                        left: 10,
                        right: 10
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: '평균등급',
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 14,
                                weight: '600'
                            },
                            color: '#2c3e50'
                        },
                        // 1~5 눈금과 격자가 정확히 보이도록 범위를 고정
                        min: 1.0,
                        max: 5.0,
                        reverse: true,
                        ticks: {
                            stepSize: 0.5,
                            callback: function(value) {
                                const roundedValue = Math.round(value * 10) / 10;
                                if (roundedValue >= 1.0 && roundedValue <= 5.0 && (roundedValue * 2) % 1 === 0) {
                                    return roundedValue.toFixed(1);
                                }
                                return '';
                            },
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 12
                            },
                            color: '#5a6c7d'
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.08)',
                            lineWidth: 1
                        }
                    },
                    y: {
                        type: 'linear',
                        display: false,
                        position: 'left',
                        min: 0,
                        max: 1
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        min: 0,
                        max: 100,
                        title: {
                            display: true,
                            text: '누적 비율 (%)',
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 14,
                                weight: '600'
                            },
                            color: '#e74c3c'
                        },
                        ticks: {
                            stepSize: 20,
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 12
                            },
                            color: '#e74c3c',
                            callback: function(value) {
                                return value + '%';
                            }
                        },
                        grid: {
                            display: false
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: {
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 13,
                                weight: '500'
                            },
                            color: '#2c3e50',
                            usePointStyle: true,
                            padding: 20
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(44, 62, 80, 0.95)',
                        titleColor: '#ffffff',
                        bodyColor: '#ffffff',
                        borderColor: 'rgba(52, 152, 219, 0.8)',
                        borderWidth: 1,
                        cornerRadius: 8,
                        displayColors: true,
                        titleFont: {
                            family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                            size: 14,
                            weight: '600'
                        },
                        bodyFont: {
                            family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                            size: 13
                        },
                        callbacks: {
                            title: function(context) {
                                if (context[0].datasetIndex === 0) {
                                    // 선 그래프 (누적 비율)
                                    return `평균등급 ${context[0].parsed.x.toFixed(1)} 이하`;
                                } else {
                                    // 산점도 (학생)
                                    const student = context[0].raw.student;
                                    return `${student.name}`;
                                }
                            },
                            label: function(context) {
                                if (context.datasetIndex === 0) {
                                    // 선 그래프 (누적 비율)
                                    return `${context.parsed.y.toFixed(1)}% : ${context.parsed.x.toFixed(1)}등급`;
                                } else {
                                    // 산점도 (학생)
                                    return `평균등급: ${context.parsed.x.toFixed(2)}`;
                                }
                            }
                        }
                    }
                },
                interaction: {
                    intersect: false,
                    mode: 'nearest'
                },
                animation: {
                    duration: 1000,
                    easing: 'easeOutCubic'
                }
            }
        });
    }

    createGradeDistributionChart(students) {
        const ctx = document.getElementById('barChart').getContext('2d');
        
        // 기존 차트가 있다면 파괴
        if (this.barChart) {
            this.barChart.destroy();
        }

        // 등급 구간별 분류
        const intervals = [
            { label: '1.0~1.5미만', min: 1.0, max: 1.5, count: 0 },
            { label: '1.5~2.0미만', min: 1.5, max: 2.0, count: 0 },
            { label: '2.0~2.5미만', min: 2.0, max: 2.5, count: 0 },
            { label: '2.5~3.0미만', min: 2.5, max: 3.0, count: 0 },
            { label: '3.0~3.5미만', min: 3.0, max: 3.5, count: 0 },
            { label: '3.5~4.0미만', min: 3.5, max: 4.0, count: 0 },
            { label: '4.0~4.5미만', min: 4.0, max: 4.5, count: 0 },
            { label: '4.5~5.0', min: 4.5, max: 5.0, count: 0 }
        ];

        students.forEach(student => {
            const grade = student.weightedAverageGrade;
            intervals.forEach(interval => {
                if (grade >= interval.min && (grade < interval.max || (interval.max === 5.0 && grade <= interval.max))) {
                    interval.count++;
                }
            });
        });

        // 누적 비율 계산 (1등급부터 누적 = 상위권부터 누적)
        const totalStudents = students.length;
        let cumulative = 0;
        const cumulativePercentages = intervals.map(interval => {
            cumulative += interval.count;
            return totalStudents > 0 ? (cumulative / totalStudents) * 100 : 0;
        });

        this.barChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: intervals.map(interval => interval.label),
                datasets: [{
                    label: '학생 수',
                    data: intervals.map(interval => interval.count),
                    backgroundColor: [
                        'rgba(26, 188, 156, 0.85)',  // 1.0-1.5 민트 그린
                        'rgba(52, 152, 219, 0.85)',  // 1.5-2.0 블루
                        'rgba(155, 89, 182, 0.85)',  // 2.0-2.5 퍼플
                        'rgba(241, 196, 15, 0.85)',  // 2.5-3.0 옐로우
                        'rgba(230, 126, 34, 0.85)',  // 3.0-3.5 오렌지
                        'rgba(231, 76, 60, 0.85)',   // 3.5-4.0 레드
                        'rgba(189, 195, 199, 0.85)', // 4.0-4.5 라이트 그레이
                        'rgba(127, 140, 141, 0.85)'  // 4.5-5.0 다크 그레이
                    ],
                    borderColor: [
                        'rgba(26, 188, 156, 1)',
                        'rgba(52, 152, 219, 1)',
                        'rgba(155, 89, 182, 1)',
                        'rgba(241, 196, 15, 1)',
                        'rgba(230, 126, 34, 1)',
                        'rgba(231, 76, 60, 1)',
                        'rgba(189, 195, 199, 1)',
                        'rgba(127, 140, 141, 1)'
                    ],
                    borderWidth: 2,
                    borderRadius: 4,
                    borderSkipped: false,
                    yAxisID: 'y',
                    // 가장자리 막대가 잘리지 않도록 여유
                    clip: 8
                }, {
                    label: '누적 비율',
                    type: 'line',
                    data: cumulativePercentages,
                    borderColor: 'rgba(231, 76, 60, 1)',
                    backgroundColor: 'rgba(231, 76, 60, 0.1)',
                    borderWidth: 3,
                    pointBackgroundColor: 'rgba(231, 76, 60, 1)',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    pointRadius: 6,
                    pointHoverRadius: 8,
                    fill: false,
                    tension: 0.2,
                    yAxisID: 'y1',
                    // 선의 끝 점이 잘리지 않도록 여유
                    clip: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: {
                    padding: {
                        top: 20,
                        bottom: 10
                    }
                },
                scales: {
                    y: {
                        type: 'linear',
                        display: true,
                        position: 'left',
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '학생 수 (명)',
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 14,
                                weight: '600'
                            },
                            color: '#2c3e50'
                        },
                        ticks: {
                            stepSize: 1,
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 12
                            },
                            color: '#5a6c7d'
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.08)',
                            lineWidth: 1
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        min: 0,
                        max: 100,
                        title: {
                            display: true,
                            text: '누적 비율 (%)',
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 14,
                                weight: '600'
                            },
                            color: '#e74c3c'
                        },
                        ticks: {
                            stepSize: 20,
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 12
                            },
                            color: '#e74c3c',
                            callback: function(value) {
                                return value + '%';
                            }
                        },
                        grid: {
                            display: false
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: '등급 구간',
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 14,
                                weight: '600'
                            },
                            color: '#2c3e50'
                        },
                        // 첫/마지막 구간에 여백을 줘서 눈금과 막대가 잘리지 않게 함
                        offset: true,
                        ticks: {
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 12,
                                weight: '500'
                            },
                            color: '#5a6c7d',
                            maxRotation: 45,
                            minRotation: 0
                        },
                        grid: {
                            display: true,
                            color: 'rgba(0, 0, 0, 0.05)',
                            lineWidth: 1
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: {
                            font: {
                                family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                                size: 13,
                                weight: '500'
                            },
                            color: '#2c3e50',
                            usePointStyle: true,
                            padding: 20
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(44, 62, 80, 0.95)',
                        titleColor: '#ffffff',
                        bodyColor: '#ffffff',
                        borderColor: 'rgba(52, 152, 219, 0.8)',
                        borderWidth: 1,
                        cornerRadius: 8,
                        displayColors: true,
                        titleFont: {
                            family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                            size: 14,
                            weight: '600'
                        },
                        bodyFont: {
                            family: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
                            size: 13
                        },
                        callbacks: {
                            title: function(context) {
                                return `등급 구간: ${context[0].label}`;
                            },
                            label: function(context) {
                                if (context.datasetIndex === 0) {
                                    // 막대그래프 (학생 수)
                                    const total = context.dataset.data.reduce((sum, val) => sum + val, 0);
                                    const percentage = total > 0 ? ((context.parsed.y / total) * 100).toFixed(1) : 0;
                                    return `학생 수: ${context.parsed.y}명 (${percentage}%)`;
                                } else {
                                    // 선 그래프 (누적 비율)
                                    return `누적 비율: ${context.parsed.y.toFixed(1)}%`;
                                }
                            }
                        }
                    }
                },
                interaction: {
                    intersect: false,
                    mode: 'index'
                },
                animation: {
                    duration: 1200,
                    easing: 'easeOutQuart'
                }
            }
        });
    }

    displayStudentAnalysis() {
        if (!this.combinedData) return;

        this.populateStudentSelectors();
        const container = document.getElementById('studentTable');
        this.renderStudentTable(this.combinedData.students, this.combinedData.subjects, container);
    }

    populateStudentSelectors() {
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        
        // 학년 옵션 생성
        const grades = [...new Set(this.combinedData.students.map(s => s.grade))].sort();
        gradeSelect.innerHTML = '<option value="">전체</option>';
        grades.forEach(grade => {
            const option = document.createElement('option');
            option.value = grade;
            option.textContent = `${grade}학년`;
            gradeSelect.appendChild(option);
        });

        // 반 옵션 생성 (전체)
        const classes = [...new Set(this.combinedData.students.map(s => s.class))].sort();
        classSelect.innerHTML = '<option value="">전체</option>';
        classes.forEach(cls => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = `${cls}반`;
            classSelect.appendChild(option);
        });

        this.updateStudentOptions();
    }

    updateClassOptions() {
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const selectedGrade = gradeSelect.value;

        let students = this.combinedData.students;
        if (selectedGrade) {
            students = students.filter(s => s.grade == selectedGrade);
        }

        const classes = [...new Set(students.map(s => s.class))].sort();
        classSelect.innerHTML = '<option value="">전체</option>';
        classes.forEach(cls => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = `${cls}반`;
            classSelect.appendChild(option);
        });
    }

    updateStudentOptions() {
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');
        
        const selectedGrade = gradeSelect.value;
        const selectedClass = classSelect.value;
        const nameQuery = (studentNameSearch && studentNameSearch.value ? studentNameSearch.value.trim() : '');

        let students = this.combinedData.students;
        if (selectedGrade) {
            students = students.filter(s => s.grade == selectedGrade);
        }
        if (selectedClass) {
            students = students.filter(s => s.class == selectedClass);
        }
        if (nameQuery) {
            const q = nameQuery.toLowerCase();
            students = students.filter(s => (s.name && s.name.toLowerCase().includes(q)) || (s.originalNumber && String(s.originalNumber).includes(q)));
        }

        studentSelect.innerHTML = '<option value="">학생 선택</option>';
        students.forEach(student => {
            const option = document.createElement('option');
            option.value = student.number;
            option.textContent = `${student.originalNumber}번 - ${student.name}`;
            studentSelect.appendChild(option);
        });
        // 단일 매치 시 자동 선택
        const showBtn = document.getElementById('showStudentDetail');
        if (students.length === 1) {
            studentSelect.value = students[0].number;
            if (showBtn) showBtn.disabled = false;
        } else {
            if (showBtn) showBtn.disabled = !studentSelect.value;
        }
    }

    renderStudentTable(students, subjects, container) {
        container.innerHTML = '';

        if (students.length === 0) {
            container.innerHTML = '<p>학생 데이터가 없습니다.</p>';
            return;
        }

        // 학생 카드 방식으로 변경
        const studentsGrid = document.createElement('div');
        studentsGrid.className = 'students-grid';

        students.forEach(student => {
            const studentCard = document.createElement('div');
            studentCard.className = 'student-card';
            
            // 과목별 평균 백분위 계산
            const weightedAveragePercentile = this.calculateWeightedAveragePercentile(student, subjects);
            
            // 평균등급 기준 순위
            const averageGradeRank = student.averageGradeRank;
            const sameGradeCount = student.sameGradeCount;
            const totalGradedStudents = student.totalGradedStudents;
            
            // 과목별 정보를 간단하게 표시
            let subjectsHTML = '';
            let hasGradeSubjects = 0;
            
            subjects.forEach(subject => {
                const score = student.scores[subject.name];
                const achievement = student.achievements[subject.name];
                const grade = student.grades[subject.name];
                const percentile = student.percentiles[subject.name];
                
                if (score !== undefined && score !== null) {
                    const hasGrade = grade !== undefined && grade !== null && grade !== 'N/A' && !isNaN(grade);
                    if (hasGrade) hasGradeSubjects++;
                    
                    subjectsHTML += `
                        <div class="subject-row ${hasGrade ? '' : 'no-grade'}">
                            <span class="subject-name">${subject.name}</span>
                            <div class="subject-data">
                                <span class="subject-score">${score}점</span>
                                ${achievement ? `<span class="subject-achievement achievement ${achievement}">${achievement}</span>` : ''}
                                ${hasGrade ? `<span class="subject-grade">${grade}등급</span>` : ''}
                                ${hasGrade && (percentile !== undefined && percentile !== null) ? `<span class="subject-percentile">${percentile}%</span>` : ''}
                            </div>
                        </div>
                    `;
                }
            });
            
            studentCard.innerHTML = `
                <div class="student-card-header">
                    <div class="student-basic-info">
                        <h4>${student.name}</h4>
                        <span class="student-number">${student.grade}학년 ${student.class}반 ${student.originalNumber}번</span>
                    </div>
                    <div class="student-summary">
                        <div class="summary-row">
                            <div class="summary-metric-inline">
                                <span class="metric-label">평균등급</span>
                                <span class="metric-value">${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                            </div>
                            ${averageGradeRank !== null && averageGradeRank !== undefined ? `
                            <div class="summary-metric-inline">
                                <span class="metric-label">등급순위</span>
                                <span class="metric-value">${averageGradeRank}/${totalGradedStudents}위${sameGradeCount > 1 ? ` (${sameGradeCount}명)` : ''}</span>
                            </div>
                            ` : ''}
                        </div>
                        ${weightedAveragePercentile ? `
                        <div class="summary-row">
                            <div class="summary-metric-inline">
                                <span class="metric-label">과목평균백분위</span>
                                <span class="metric-value">${weightedAveragePercentile.toFixed(1)}%</span>
                            </div>
                        </div>
                        ` : ''}
                    </div>
                </div>
                <div class="student-subjects">
                    ${subjectsHTML}
                </div>
                <div class="student-card-footer">
                    <span class="grade-subjects-count">등급 산출 과목: ${hasGradeSubjects}개</span>
                    <button class="view-detail-btn" data-student-id="${student.number}">상세 보기</button>
                </div>
            `;
            
            studentsGrid.appendChild(studentCard);
        });

        container.appendChild(studentsGrid);

        // 카드 내 상세 보기 버튼 클릭 처리 (이벤트 위임)
        studentsGrid.addEventListener('click', (e) => {
            const btn = e.target.closest('.view-detail-btn');
            if (!btn) return;
            const studentId = btn.getAttribute('data-student-id');
            if (!studentId) return;

            // 선택 박스 동기화 (선택되어 있다면)
            const studentSelect = document.getElementById('studentSelect');
            if (studentSelect) {
                studentSelect.value = studentId;
            }

            const targetStudent = this.combinedData.students.find(s => s.number == studentId);
            if (!targetStudent) return;

            this.renderStudentDetail(targetStudent);
            this.switchView('detail');
        });
    }

    filterStudents(searchTerm) {
        if (!this.combinedData) return;

        const filtered = this.combinedData.students.filter(student => 
            student.number.toString().includes(searchTerm) || 
            student.name.includes(searchTerm) ||
            student.fileName.includes(searchTerm)
        );
        
        const container = document.getElementById('studentTable');
        this.renderStudentTable(filtered, this.combinedData.subjects, container);
    }

    switchTab(tabName) {
        // 탭 버튼 활성화
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // 탭 내용 표시
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`${tabName}-tab`).classList.add('active');
    }

    switchView(viewType) {
        const tableViewBtn = document.getElementById('tableViewBtn');
        const detailViewBtn = document.getElementById('detailViewBtn');
        const tableView = document.getElementById('tableView');
        const detailView = document.getElementById('detailView');

        if (viewType === 'table') {
            tableViewBtn.classList.add('active');
            detailViewBtn.classList.remove('active');
            tableView.style.display = 'block';
            detailView.style.display = 'none';
        } else {
            tableViewBtn.classList.remove('active');
            detailViewBtn.classList.add('active');
            tableView.style.display = 'none';
            detailView.style.display = 'block';
        }
    }

    showStudentDetail() {
        const studentSelect = document.getElementById('studentSelect');
        const selectedStudentId = studentSelect.value;
        
        if (!selectedStudentId) return;

        const student = this.combinedData.students.find(s => s.number == selectedStudentId);
        if (!student) return;

        this.renderStudentDetail(student);
        this.switchView('detail');
    }

    renderStudentDetail(student) {
        const container = document.getElementById('studentDetailContent');
        
        // 기존 학급 전체 인쇄 영역 완전 제거
        const classPrintArea = document.getElementById('classPrintArea');
        if (classPrintArea) {
            classPrintArea.remove();
        }
        
        // 학급 전체 인쇄 관련 클래스 제거
        const studentsTab = document.getElementById('students-tab');
        if (studentsTab) {
            studentsTab.classList.remove('only-class-print', 'print-target');
        }
        
        // 학점 가중 평균 백분위 계산
        const weightedAveragePercentile = this.calculateWeightedAveragePercentile(student, this.combinedData.subjects);
        
        // 평균등급 기준 순위
        const averageGradeRank = student.averageGradeRank;
        const sameGradeCount = student.sameGradeCount;
        const totalGradedStudents = student.totalGradedStudents;
        
        const html = `
            <div class="print-controls">
                <button class="print-btn" onclick="scoreAnalyzer.printStudentDetail('${student.name}')">프린터 출력</button>
                <button class="pdf-btn" onclick="scoreAnalyzer.generatePDF('${student.name}')">PDF 저장</button>
            </div>
            
            <div id="printArea" class="print-area">
                <div class="print-header" style="display: none;">
                    <h2>학생 성적 분석 보고서</h2>
                    <div class="print-date">생성일: ${new Date().toLocaleDateString('ko-KR')}</div>
                </div>
                
                <div class="student-detail-header">
                    <div class="student-info">
                        <h3>${student.name}</h3>
                        <div class="student-meta">
                            <span class="grade-class">${student.grade}학년 ${student.class}반 ${student.originalNumber}번</span>
                            <span class="file-info">출처: ${student.fileName}</span>
                        </div>
                    </div>
                    <div class="overall-stats">
                        <div class="stat-card">
                            <span class="stat-label">평균등급</span>
                            <span class="stat-value grade">${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                        </div>
                        <div class="stat-card">
                            <span class="stat-label">전체 학생수</span>
                            <span class="stat-value">${student.totalStudents || 'N/A'}명</span>
                        </div>
                    </div>
                </div>
                
                <div class="student-detail-content">
                    <div class="analysis-overview">
                        <div class="student-summary">
                            <div class="summary-card">
                                <div class="summary-header">
                                    <h4>학생 정보</h4>
                                </div>
                                <div class="summary-grid">
                                    <div class="summary-item">
                                        <span class="summary-label">학급</span>
                                        <span class="summary-value">${student.grade}학년 ${student.class}반 ${student.originalNumber}번</span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">평균등급</span>
                                        <span class="summary-value highlight">${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">평균등급(9등급환산)</span>
                                        <span class="summary-value orange">${student.weightedAverage9Grade ? student.weightedAverage9Grade.toFixed(2) : 'N/A'}</span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">등급 순위</span>
                                        <span class="summary-value highlight">${averageGradeRank !== null && averageGradeRank !== undefined ? `${averageGradeRank}/${totalGradedStudents}위` + (sameGradeCount > 1 ? ` (${sameGradeCount}명)` : '') : 'N/A'}</span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">과목평균 백분위</span>
                                        <span class="summary-value highlight">${weightedAveragePercentile ? weightedAveragePercentile.toFixed(1) + '%' : 'N/A'}</span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">전체 학생수</span>
                                        <span class="summary-value">${student.totalStudents || 'N/A'}명</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="chart-container">
                            <h4>과목별 등급</h4>
                            <canvas id="studentPercentileChart" width="400" height="400"></canvas>
                        </div>
                    </div>
                    
                    <div class="subject-details">
                        <h4>과목별 상세 분석</h4>
                        <div class="subject-cards">
                            ${this.renderSubjectCards(student)}
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        container.innerHTML = html;
        
        // 레이더 차트 생성
        setTimeout(() => {
            this.createStudentPercentileChart(student);
        }, 100);
    }

    // 학급 전체 인쇄용: 개별 학생과 완전히 동일한 HTML 구조
    buildStudentDetailHTMLForPrint(student, canvasId) {
        const weightedAveragePercentile = this.calculateWeightedAveragePercentile(student, this.combinedData.subjects);
        const averageGradeRank = student.averageGradeRank;
        const sameGradeCount = student.sameGradeCount;
        const totalGradedStudents = student.totalGradedStudents;
        return `
            <div class="student-print-page">
                <div id="printArea-${canvasId}" class="print-area">
                    <div class="print-header" style="display: none;">
                        <h2>학생 성적 분석 보고서</h2>
                        <div class="print-date">생성일: ${new Date().toLocaleDateString('ko-KR')}</div>
                    </div>
                    
                    <div class="student-detail-header">
                        <div class="student-info">
                            <h3>${student.name}</h3>
                            <div class="student-meta">
                                <span class="grade-class">${student.grade}학년 ${student.class}반 ${student.originalNumber}번</span>
                                <span class="file-info">출처: ${student.fileName}</span>
                            </div>
                        </div>
                        <div class="overall-stats">
                            <div class="stat-card">
                                <span class="stat-label">평균등급</span>
                                <span class="stat-value grade">${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                            </div>
                            <div class="stat-card">
                                <span class="stat-label">전체 학생수</span>
                                <span class="stat-value">${student.totalStudents || 'N/A'}명</span>
                            </div>
                        </div>
                    </div>
                    
                    <div class="student-detail-content">
                        <div class="analysis-overview">
                            <div class="student-summary">
                                <div class="summary-card">
                                    <div class="summary-header">
                                        <h4>학생 정보</h4>
                                    </div>
                                    <div class="summary-grid">
                                        <div class="summary-item">
                                            <span class="summary-label">학급</span>
                                            <span class="summary-value">${student.grade}학년 ${student.class}반 ${student.originalNumber}번</span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">평균등급</span>
                                            <span class="summary-value highlight">${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">평균등급(9등급환산)</span>
                                            <span class="summary-value orange">${student.weightedAverage9Grade ? student.weightedAverage9Grade.toFixed(2) : 'N/A'}</span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">등급 순위</span>
                                            <span class="summary-value highlight">${averageGradeRank !== null && averageGradeRank !== undefined ? `${averageGradeRank}/${totalGradedStudents}위` + (sameGradeCount > 1 ? ` (${sameGradeCount}명)` : '') : 'N/A'}</span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">과목평균 백분위</span>
                                            <span class="summary-value highlight">${weightedAveragePercentile ? weightedAveragePercentile.toFixed(1) + '%' : 'N/A'}</span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">전체 학생수</span>
                                            <span class="summary-value">${student.totalStudents || 'N/A'}명</span>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="chart-container">
                                <h4>과목별 등급</h4>
                                <canvas id="${canvasId}" width="400" height="400"></canvas>
                            </div>
                        </div>
                        
                        <div class="subject-details">
                            <h4>과목별 상세 분석</h4>
                            <div class="subject-cards">
                                ${this.renderSubjectCards(student)}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    // 다중 생성용 차트 (개별 PDF와 동일한 설정)
    createStudentPercentileChartFor(canvas, student) {
        if (!canvas) return null;
        const subjects = this.combinedData.subjects.filter(subject => {
            const grade = student.grades[subject.name];
            return grade !== undefined && grade !== null && grade !== 'N/A' && !isNaN(grade);
        });
        if (subjects.length === 0) return null;
        const labels = subjects.map(subject => subject.name);
        const gradeData = subjects.map(subject => {
            const grade = student.grades[subject.name];
            return grade ? (6 - grade) : 0;
        });
        return new Chart(canvas, {
            type: 'radar',
            plugins: [ChartDataLabels],
            data: {
                labels,
                datasets: [{
                    label: '등급',
                    data: gradeData,
                    backgroundColor: 'rgba(52, 152, 219, 0.2)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 2,
                    pointBackgroundColor: 'rgba(52, 152, 219, 1)',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                animation: {
                    duration: 0
                },
                interaction: {
                    intersect: false
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        enabled: false
                    },
                    datalabels: {
                        display: true,
                        formatter: function(value, context) {
                            const subjectIndex = context.dataIndex;
                            const subject = subjects[subjectIndex];
                            const grade = student.grades[subject.name];
                            return `${grade}등급`;
                        },
                        color: '#2c3e50',
                        backgroundColor: 'rgba(255, 255, 255, 0.9)',
                        borderColor: '#dee2e6',
                        borderWidth: 1,
                        borderRadius: 4,
                        padding: 4,
                        font: {
                            size: 11,
                            weight: '500'
                        }
                    }
                },
                scales: {
                    r: {
                        beginAtZero: true,
                        max: 5,
                        min: 0,
                        ticks: {
                            stepSize: 1,
                            color: '#5a6c7d',
                            callback: function(value) {
                                if (value === 0) return '';
                                return `${6 - value}등급`;
                            }
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.1)'
                        },
                        angleLines: {
                            color: 'rgba(0, 0, 0, 0.1)'
                        },
                        pointLabels: {
                            font: {
                                size: 12,
                                weight: '500'
                            },
                            color: '#2c3e50'
                        }
                    }
                }
            }
        });
    }


    // 학급 전체 PDF
    async generateSelectedClassPDF() {
        // 필요 변수는 try 외부에 선언하여 예외 처리에서 접근 가능하도록 함
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const grade = gradeSelect.value;
        const cls = classSelect.value;
        let students = [];
        try {
            if (!grade || !cls) {
                alert('학년과 반을 선택해 주세요.');
                return;
            }
            students = this.combinedData.students.filter(s => String(s.grade) === String(grade) && String(s.class) === String(cls));
            if (students.length === 0) {
                alert('선택한 학급의 학생이 없습니다.');
                return;
            }

            const { jsPDF } = window.jspdf;
            // 메모리 사용을 줄이기 위해 압축 활성화
            const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4', compress: true });
            const pdfWidth = 210, pdfHeight = 297;
            const maxImgWidth = pdfWidth - 20; // 10mm 여백
            const maxImgHeight = pdfHeight - 20; // 상하 10mm 여백

            // 임시 캡처 컨테이너
            const temp = document.createElement('div');
            temp.style.position = 'fixed';
            temp.style.left = '-10000px';
            temp.style.top = '0';
            document.body.appendChild(temp);

            for (let i = 0; i < students.length; i++) {
                const student = students[i];
                const canvasId = `pdfRadar-${student.grade}-${student.class}-${student.number}-${i}`;
                temp.innerHTML = this.buildStudentDetailHTMLForPrint(student, canvasId);
                // 차트 렌더
                await new Promise(r => setTimeout(r, 50));
                const canvas = document.getElementById(canvasId);
                const chartInstance = canvas ? this.createStudentPercentileChartFor(canvas, student) : null;
                await new Promise(r => setTimeout(r, 200));

                const element = temp.firstElementChild;
                // 캔버스 스케일을 낮추고 JPEG로 변환하여 용량 축소
                const canvasImg = await html2canvas(element, { scale: 1.3, backgroundColor: '#ffffff', useCORS: true, allowTaint: true });
                const imgData = canvasImg.toDataURL('image/jpeg', 0.82);
                const aspect = canvasImg.width / canvasImg.height;
                let drawWidth = maxImgWidth;
                let drawHeight = drawWidth / aspect;
                if (drawHeight > maxImgHeight) { drawHeight = maxImgHeight; drawWidth = drawHeight * aspect; }
                const x = (pdfWidth - drawWidth) / 2;
                const y = (pdfHeight - drawHeight) / 2;

                if (i > 0) pdf.addPage();
                pdf.addImage(imgData, 'JPEG', x, y, drawWidth, drawHeight);

                // 차트 메모리 해제
                if (chartInstance && typeof chartInstance.destroy === 'function') {
                    try { chartInstance.destroy(); } catch (_) {}
                }
            }

            document.body.removeChild(temp);
            const fileName = `${grade}학년_${cls}반_학생성적_${new Date().toISOString().split('T')[0]}.pdf`;
            pdf.save(fileName);
        } catch (err) {
            console.error('학급 전체 PDF 생성 오류:', err);
            // 문자열 길이 초과 등으로 실패하는 경우, 파일을 여러 개로 나눠 저장을 시도
            const isLenErr = err && (err.name === 'RangeError' || String(err.message || '').includes('Invalid string length'));
            if (isLenErr && students && students.length > 0) {
                try {
                    const chunkSize = 12; // 용량 방지를 위한 페이지 분할 크기
                    const totalParts = Math.ceil(students.length / chunkSize);
                    for (let part = 0; part < totalParts; part++) {
                        const start = part * chunkSize;
                        const end = Math.min(students.length, start + chunkSize);
                        const { jsPDF } = window.jspdf;
                        const partPdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4', compress: true });
                        const pdfWidth = 210, pdfHeight = 297;
                        const maxImgWidth = pdfWidth - 20;
                        const maxImgHeight = pdfHeight - 20;

                        const temp = document.createElement('div');
                        temp.style.position = 'fixed';
                        temp.style.left = '-10000px';
                        temp.style.top = '0';
                        document.body.appendChild(temp);

                        for (let i = start; i < end; i++) {
                            const student = students[i];
                            const canvasId = `pdfRadar-${student.grade}-${student.class}-${student.number}-${i}`;
                            temp.innerHTML = this.buildStudentDetailHTMLForPrint(student, canvasId);
                            await new Promise(r => setTimeout(r, 50));
                            const canvas = document.getElementById(canvasId);
                            const chartInstance = canvas ? this.createStudentPercentileChartFor(canvas, student) : null;
                            await new Promise(r => setTimeout(r, 200));

                            const element = temp.firstElementChild;
                            const canvasImg = await html2canvas(element, { scale: 1.3, backgroundColor: '#ffffff', useCORS: true, allowTaint: true });
                            const imgData = canvasImg.toDataURL('image/jpeg', 0.82);
                            const aspect = canvasImg.width / canvasImg.height;
                            let drawWidth = maxImgWidth;
                            let drawHeight = drawWidth / aspect;
                            if (drawHeight > maxImgHeight) { drawHeight = maxImgHeight; drawWidth = drawHeight * aspect; }
                            const x = (pdfWidth - drawWidth) / 2;
                            const y = (pdfHeight - drawHeight) / 2;

                            if (i > start) partPdf.addPage();
                            partPdf.addImage(imgData, 'JPEG', x, y, drawWidth, drawHeight);

                            if (chartInstance && typeof chartInstance.destroy === 'function') {
                                try { chartInstance.destroy(); } catch (_) {}
                            }
                        }

                        document.body.removeChild(temp);
                        const partName = `${grade}학년_${cls}반_학생성적_${new Date().toISOString().split('T')[0]}_part${part + 1}-of-${totalParts}.pdf`;
                        partPdf.save(partName);
                    }
                    alert('PDF가 용량 문제로 여러 개의 파일로 분할 저장되었습니다.');
                    return;
                } catch (fallbackErr) {
                    console.error('분할 저장 시도 중 오류:', fallbackErr);
                }
            }
            alert('학급 전체 PDF 생성 중 오류가 발생했습니다: ' + (err && err.message ? err.message : String(err)));
        }
    }

    renderSubjectCards(student) {
        return this.combinedData.subjects.map(subject => {
            const score = student.scores[subject.name] || 0;
            const achievement = student.achievements[subject.name] || 'N/A';
            const grade = student.grades ? student.grades[subject.name] : undefined;
            const rank = student.ranks ? student.ranks[subject.name] || 'N/A' : 'N/A';
            const percentile = student.percentiles ? student.percentiles[subject.name] || 0 : 0;
            
            // 등급이 있는지 확인
            const hasGrade = grade !== undefined && grade !== null && grade !== 'N/A' && !isNaN(grade);
            
            // 백분위에 따른 색상 결정 (등급이 있는 경우만)
            let percentileClass = 'low';
            if (hasGrade && percentile >= 80) percentileClass = 'excellent';
            else if (hasGrade && percentile >= 60) percentileClass = 'good';
            else if (hasGrade && percentile >= 40) percentileClass = 'average';
            
            if (hasGrade) {
                // 등급이 있는 과목: 모든 정보 표시
                return `
                    <div class="subject-card">
                        <div class="subject-header">
                            <h5>${subject.name}</h5>
                            <span class="credits">${subject.credits}학점</span>
                        </div>
                        <div class="subject-metrics">
                            <div class="metric">
                                <span class="metric-label">점수</span>
                                <span class="metric-value">${score}점</span>
                                <span class="metric-average">(평균: ${subject.average ? subject.average.toFixed(1) : 'N/A'}점)</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">성취도</span>
                                <span class="metric-value achievement ${achievement}">${achievement}</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">등급</span>
                                <span class="metric-value">${grade}등급</span>
                            </div>
                        </div>
                        <div class="subject-metrics">
                            <div class="metric">
                                <span class="metric-label">석차</span>
                                <span class="metric-value">${rank}위</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">백분위</span>
                                <span class="metric-value percentile ${percentileClass}">${percentile}%</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">등급(9등급환산)</span>
                                <span class="metric-value orange">${this.convertPercentileTo9Grade(percentile) || 'N/A'}등급</span>
                            </div>
                        </div>
                        <div class="percentile-bar">
                            <div class="percentile-fill ${percentileClass}" style="width: ${percentile}%"></div>
                        </div>
                    </div>
                `;
            } else {
                // 등급이 없는 과목: 점수, 평균, 성취도만 표시
                return `
                    <div class="subject-card no-grade">
                        <div class="subject-header">
                            <h5>${subject.name}</h5>
                            <span class="credits">${subject.credits}학점</span>
                        </div>
                        <div class="subject-metrics simple">
                            <div class="metric">
                                <span class="metric-label">점수</span>
                                <span class="metric-value">${score}점</span>
                                <span class="metric-average">(평균: ${subject.average ? subject.average.toFixed(1) : 'N/A'}점)</span>
                            </div>
                            <div class="metric">
                                <span class="metric-label">성취도</span>
                                <span class="metric-value achievement ${achievement}">${achievement}</span>
                            </div>
                        </div>
                        <div class="no-grade-notice">
                            <span>등급 산출 대상 과목이 아닙니다</span>
                        </div>
                    </div>
                `;
            }
        }).join('');
    }

    createStudentPercentileChart(student) {
        const ctx = document.getElementById('studentPercentileChart');
        if (!ctx) return;
        
        // 기존 차트 제거
        if (this.studentPercentileChart) {
            this.studentPercentileChart.destroy();
        }

        // 등급이 있는 과목만 필터링
        const subjects = this.combinedData.subjects.filter(subject => {
            const grade = student.grades[subject.name];
            return grade !== undefined && grade !== null && grade !== 'N/A' && !isNaN(grade);
        });

        if (subjects.length === 0) {
            ctx.parentElement.style.display = 'none';
            return;
        }

        ctx.parentElement.style.display = 'block';
        const labels = subjects.map(subject => subject.name);
        const gradeData = subjects.map(subject => {
            const grade = student.grades[subject.name];
            // 등급을 역순으로 변환 (1등급=5, 2등급=4, ..., 5등급=1)하여 차트에서 높게 표시
            return grade ? (6 - grade) : 0;
        });
        
        this.studentPercentileChart = new Chart(ctx, {
            type: 'radar',
            plugins: [ChartDataLabels],
            data: {
                labels: labels,
                datasets: [{
                    label: '등급',
                    data: gradeData,
                    backgroundColor: 'rgba(52, 152, 219, 0.2)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 2,
                    pointBackgroundColor: 'rgba(52, 152, 219, 1)',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                interaction: {
                    intersect: false
                },
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        backgroundColor: 'rgba(44, 62, 80, 0.95)',
                        titleColor: '#ffffff',
                        bodyColor: '#ffffff',
                        callbacks: {
                            label: function(context) {
                                const subjectName = context.label;
                                const gradeValue = context.parsed.r;
                                // 역순으로 변환된 값을 다시 등급으로 변환
                                const grade = gradeValue > 0 ? (6 - gradeValue) : 'N/A';
                                return `${grade}등급`;
                            }
                        }
                    },
                    datalabels: {
                        display: true,
                        color: '#2c3e50',
                        backgroundColor: 'rgba(255, 255, 255, 0.8)',
                        borderColor: '#dee2e6',
                        borderWidth: 1,
                        borderRadius: 4,
                        padding: {
                            top: 4,
                            bottom: 4,
                            left: 6,
                            right: 6
                        },
                        font: {
                            size: 11,
                            weight: 'bold'
                        },
                        formatter: function(value, context) {
                            const subjectIndex = context.dataIndex;
                            const grade = subjects[subjectIndex] ? student.grades[subjects[subjectIndex].name] : 'N/A';
                            return `${grade}등급`;
                        },
                        anchor: 'end',
                        align: 'top',
                        offset: 10,
                        textAlign: 'center'
                    }
                },
                scales: {
                    r: {
                        beginAtZero: true,
                        max: 5,
                        min: 0,
                        ticks: {
                            stepSize: 1,
                            font: {
                                size: 12
                            },
                            color: '#5a6c7d',
                            callback: function(value) {
                                // 역순으로 표시 (5가 1등급, 1이 5등급)
                                if (value === 0) return '';
                                return `${6 - value}등급`;
                            }
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.1)'
                        },
                        angleLines: {
                            color: 'rgba(0, 0, 0, 0.1)'
                        },
                        pointLabels: {
                            font: {
                                size: 12,
                                weight: '500'
                            },
                            color: '#2c3e50'
                        }
                    }
                }
            }
        });
    }

    // 프린터 출력 기능
    printStudentDetail(studentName) {
        try {
            // 인쇄 전용 클래스 설정
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('print-target');
            });
            document.getElementById('students-tab').classList.add('print-target');
            
            // 인쇄 영역을 한 페이지에 맞게 스케일
            const printArea = document.getElementById('printArea') || document.getElementById('studentDetailContent');
            if (printArea) {
                // mm -> px 변환요소 생성
                const mm = document.createElement('div');
                mm.style.width = '1mm';
                mm.style.height = '1mm';
                mm.style.position = 'absolute';
                mm.style.visibility = 'hidden';
                document.body.appendChild(mm);
                const pxPerMM = mm.getBoundingClientRect().width || 3.78; // fallback 96dpi 기준
                document.body.removeChild(mm);

                const printableWidthPx = (210 - 20) * pxPerMM;  // 10mm 좌우 여백
                const printableHeightPx = (297 - 20) * pxPerMM; // 10mm 상하 여백
                const rect = printArea.getBoundingClientRect();
                const scale = Math.min(printableWidthPx / rect.width, printableHeightPx / rect.height, 1);
                printArea.style.setProperty('--page-scale', String(scale));
                printArea.classList.add('apply-print-scale');

                const cleanup = () => {
                    printArea.classList.remove('apply-print-scale');
                    printArea.style.removeProperty('--page-scale');
                    window.removeEventListener('afterprint', cleanup);
                };
                window.addEventListener('afterprint', cleanup);
                
                // 인쇄 실행
                window.print();
                
                // 일부 브라우저용 안전망
                setTimeout(() => cleanup(), 1000);
            } else {
                // 기본 인쇄
                window.print();
            }
            
            // 인쇄 후 별도 처리 없음
            
        } catch (error) {
            console.error('프린터 출력 중 오류:', error);
            alert('프린터 출력 중 오류가 발생했습니다: ' + error.message);
        }
    }

    // PDF 생성 기능
    async generatePDF(studentName) {
        try {
            // 인쇄 전용 클래스 설정
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('print-target');
            });
            document.getElementById('students-tab').classList.add('print-target');
            
            // 잠시 기다려 레이아웃 적용
            await new Promise(resolve => setTimeout(resolve, 100));
            
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF('p', 'mm', 'a4');
            
            // PDF에 포함할 요소 선택 (차트 제외)
            const element = document.getElementById('printArea');
            if (!element) {
                alert('PDF 생성할 내용을 찾을 수 없습니다.');
                return;
            }

            // html2canvas로 요소를 캡처
            const canvas = await html2canvas(element, {
                scale: 2,
                backgroundColor: '#ffffff',
                width: element.scrollWidth,
                height: element.scrollHeight,
                useCORS: true,
                allowTaint: true
            });

            const imgData = canvas.toDataURL('image/png');
            
            // PDF 크기 계산 (한 페이지에 맞춤)
            const pdfWidth = 210; // A4 width in mm
            const pdfHeight = 297; // A4 height in mm
            const maxImgWidth = pdfWidth - 20;  // 좌우 여백 합 20mm
            const maxImgHeight = pdfHeight - 60; // 상단 제목/정보 여백 60mm
            const imgAspect = canvas.width / canvas.height;
            let drawWidth = maxImgWidth;
            let drawHeight = drawWidth / imgAspect;
            if (drawHeight > maxImgHeight) {
                drawHeight = maxImgHeight;
                drawWidth = drawHeight * imgAspect;
            }

            // 이미지가 한 페이지에 들어가는지 확인
            // 한 페이지에 맞춰 중앙 정렬하여 배치 (상하 여백 10mm 기준)
            const x = (pdfWidth - drawWidth) / 2;
            const y = 10 + (maxImgHeight - drawHeight) / 2;
            pdf.addImage(imgData, 'PNG', x, y, drawWidth, drawHeight);

            // PDF 다운로드
            const fileName = `${studentName}_성적분석_${new Date().toISOString().split('T')[0]}.pdf`;
            pdf.save(fileName);

        } catch (error) {
            console.error('PDF 생성 중 오류:', error);
            alert('PDF 생성 중 오류가 발생했습니다: ' + error.message);
        }
    }


    showLoading() {
        document.getElementById('loading').style.display = 'block';
        document.getElementById('results').style.display = 'none';
        this.hideError();
    }

    hideLoading() {
        document.getElementById('loading').style.display = 'none';
    }

    showError(message) {
        const errorDiv = document.getElementById('error');
        errorDiv.textContent = message;
        errorDiv.style.display = 'block';
    }

    hideError() {
        document.getElementById('error').style.display = 'none';
    }
}

// 전역 변수로 선언
let scoreAnalyzer;

// 페이지 로드 시 분석기 초기화
document.addEventListener('DOMContentLoaded', () => {
    scoreAnalyzer = new ScoreAnalyzer();
});
