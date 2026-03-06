class ScoreAnalyzer {
    constructor() {
        this.filesData = new Map(); // 파일명 -> 분석 데이터 매핑
        this.combinedData = null; // 통합된 분석 데이터
        this.selectedFiles = null; // 사용자가 선택/드롭한 파일 목록
        this.subjectGroups = null; // 교과(군) 매핑 데이터
        this.subjectGroupsReady = this.loadSubjectGroups(); // 교과(군) 데이터 로드
        this.handleStudentDetailKeydown = this.handleStudentDetailKeydown.bind(this);
        this.initializeEventListeners();

        // If the page provides preloaded analysis data, render directly
        if (window.PRELOADED_DATA) {
            this.initializePreloadedView();
        }
    }

    async initializePreloadedView() {
        try {
            await this.subjectGroupsReady;
            this.combinedData = window.PRELOADED_DATA;
            // 저장된 HTML에서는 업로드 섹션 숨기기
            if (window.PRELOADED_HIDE_UPLOAD) {
                const uploadSection = document.querySelector('.upload-section');
                if (uploadSection) uploadSection.style.display = 'none';
            }
            const results = document.getElementById('results');
            if (results) results.style.display = 'block';
            this.displayResults();
            this.applyPreloadedUiState();
            const exportCsvBtn = document.getElementById('exportCsvBtn');
            const exportHtmlBtn = document.getElementById('exportHtmlBtn');
            if (exportCsvBtn) exportCsvBtn.disabled = false;
            if (exportHtmlBtn) exportHtmlBtn.disabled = false;
        } catch (e) {
            console.error('PRELOADED_DATA 처리 중 오류:', e);
        }
    }

    setIntroSectionVisible(visible) {
        const container = document.querySelector('.container');
        if (!container) return;
        container.classList.toggle('post-analysis', !visible);
    }

    getCurrentUiState() {
        const activeTabBtn = document.querySelector('.tab-btn.active');
        const detailViewBtn = document.getElementById('detailViewBtn');

        return {
            activeTab: activeTabBtn ? activeTabBtn.dataset.tab : 'grade-analysis',
            activeView: detailViewBtn && detailViewBtn.classList.contains('active') ? 'detail' : 'table',
            selectedGrade: document.getElementById('gradeSelect')?.value || '',
            selectedClass: document.getElementById('classSelect')?.value || '',
            selectedStudent: document.getElementById('studentSelect')?.value || '',
            studentNameSearch: document.getElementById('studentNameSearch')?.value || '',
            studentSearch: document.getElementById('studentSearch')?.value || ''
        };
    }

    applyPreloadedUiState() {
        const state = window.PRELOADED_UI_STATE;
        if (!state || !this.combinedData) return;

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');
        const studentSearch = document.getElementById('studentSearch');
        const showStudentDetail = document.getElementById('showStudentDetail');

        if (gradeSelect) gradeSelect.value = state.selectedGrade || '';
        this.updateClassOptions();

        if (classSelect) classSelect.value = state.selectedClass || '';
        if (studentNameSearch) studentNameSearch.value = state.studentNameSearch || '';
        this.updateStudentOptions();

        if (studentSelect && state.selectedStudent) {
            studentSelect.value = String(state.selectedStudent);
        }
        if (studentSearch) {
            studentSearch.value = state.studentSearch || '';
        }
        if (showStudentDetail && studentSelect) {
            showStudentDetail.disabled = !studentSelect.value;
        }

        this.filterStudentTable();

        if (state.activeTab && document.querySelector(`[data-tab="${state.activeTab}"]`)) {
            this.switchTab(state.activeTab);
        }

        if (state.activeView === 'detail' && studentSelect && studentSelect.value) {
            const targetStudent = this.combinedData.students.find(
                student => String(student.number) === String(studentSelect.value)
            );
            if (targetStudent) {
                this.renderStudentDetail(targetStudent);
                this.switchView('detail');
            }
        } else if (state.activeView === 'table') {
            this.switchView('table');
        }
    }

    // 교과(군) 매핑 데이터 로드
    async loadSubjectGroups() {
        if (window.PRELOADED_SUBJECT_GROUPS) {
            this.subjectGroups = window.PRELOADED_SUBJECT_GROUPS;
            return this.subjectGroups;
        }

        try {
            const response = await fetch('subjectGroups.json');
            if (response.ok) {
                this.subjectGroups = await response.json();
                console.log('교과(군) 매핑 데이터 로드 완료');
            } else {
                console.warn('subjectGroups.json 파일을 찾을 수 없습니다. 기본 매핑을 사용합니다.');
                this.setDefaultSubjectGroups();
            }
        } catch (error) {
            console.warn('교과(군) 매핑 데이터 로드 실패:', error);
            this.setDefaultSubjectGroups();
        }

        return this.subjectGroups;
    }

    // 기본 교과(군) 매핑 설정 (JSON 로드 실패 시)
    setDefaultSubjectGroups() {
        this.subjectGroups = {
            groups: {
                "국어": { keywords: ["국어", "화법", "독서", "문학", "언어", "작문", "매체"], color: "#e74c3c", order: 1 },
                "수학": { keywords: ["수학", "대수", "미적분", "확률", "통계", "기하"], color: "#3498db", order: 2 },
                "영어": { keywords: ["영어", "English"], color: "#2ecc71", order: 3 },
                "사회": { keywords: ["사회", "역사", "지리", "윤리", "정치", "경제", "법", "세계사", "동아시아", "시민"], color: "#f39c12", order: 4 },
                "과학": { keywords: ["과학", "물리", "화학", "생명", "지구", "탐구실험", "생물"], color: "#9b59b6", order: 5 },
                "기타": { keywords: [], color: "#95a5a6", order: 6 }
            },
            exactMatch: {
                "한국사1": "사회", "한국사2": "사회",
                "통합사회1": "사회", "통합사회2": "사회",
                "통합과학1": "과학", "통합과학2": "과학",
                "정보": "기타", "기술가정": "기타", "음악": "기타", "미술": "기타", "체육": "기타"
            }
        };
    }

    // 과목명을 교과(군)으로 매핑
    getSubjectGroup(subjectName) {
        if (!this.subjectGroups) {
            return "기타";
        }

        // 1. 정확히 일치하는 항목 먼저 확인
        if (this.subjectGroups.exactMatch && this.subjectGroups.exactMatch[subjectName]) {
            return this.subjectGroups.exactMatch[subjectName];
        }

        // 2. 키워드 기반 매핑
        for (const [groupName, groupData] of Object.entries(this.subjectGroups.groups)) {
            if (groupName === "기타") continue; // 기타는 마지막에 처리
            for (const keyword of groupData.keywords) {
                if (subjectName.includes(keyword)) {
                    return groupName;
                }
            }
        }

        // 3. 매칭되지 않으면 기타
        return "기타";
    }

    getSubjectColumnLabel(subject) {
        if (!subject || !subject.name) return '';
        return `${this.getSubjectGroup(subject.name)}_${subject.name}`;
    }

    // 학생의 교과(군)별 평균 등급 계산
    calculateGroupGrades(student) {
        const groupData = {};

        // 교과군별로 데이터 초기화
        if (this.subjectGroups && this.subjectGroups.groups) {
            for (const groupName of Object.keys(this.subjectGroups.groups)) {
                groupData[groupName] = {
                    totalGradePoints: 0,
                    totalCredits: 0,
                    subjects: []
                };
            }
        }

        // 각 과목을 교과군에 할당
        this.combinedData.subjects.forEach(subject => {
            const grade = student.grades[subject.name];
            if (grade !== undefined && grade !== null && !isNaN(grade)) {
                const groupName = this.getSubjectGroup(subject.name);
                if (!groupData[groupName]) {
                    groupData[groupName] = { totalGradePoints: 0, totalCredits: 0, subjects: [] };
                }
                groupData[groupName].totalGradePoints += grade * subject.credits;
                groupData[groupName].totalCredits += subject.credits;
                groupData[groupName].subjects.push({
                    name: subject.name,
                    grade: grade,
                    credits: subject.credits
                });
            }
        });

        // 교과군별 평균 등급 계산
        const result = {};
        for (const [groupName, data] of Object.entries(groupData)) {
            if (data.totalCredits > 0) {
                result[groupName] = {
                    averageGrade: data.totalGradePoints / data.totalCredits,
                    totalCredits: data.totalCredits,
                    subjects: data.subjects,
                    color: this.subjectGroups?.groups?.[groupName]?.color || "#95a5a6",
                    order: this.subjectGroups?.groups?.[groupName]?.order || 99
                };
            }
        }

        return result;
    }

    initializeEventListeners() {
        const fileInput = document.getElementById('excelFiles');
        const analyzeBtn = document.getElementById('analyzeBtn');
        const exportCsvBtn = document.getElementById('exportCsvBtn');
        const exportHtmlBtn = document.getElementById('exportHtmlBtn');
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

        document.addEventListener('keydown', this.handleStudentDetailKeydown);

        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length > 0) {
                this.selectedFiles = files;
                this.displayFileList(files);
                analyzeBtn.disabled = false;
                this.hideError();
            }
        });

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

        if (analyzeBtn) analyzeBtn.addEventListener('click', () => { this.analyzeFiles(); });

        if (exportCsvBtn) exportCsvBtn.addEventListener('click', () => { this.exportToCSV(); });

        if (exportHtmlBtn) exportHtmlBtn.addEventListener('click', () => { this.showHtmlExportOptionsModal(); });

        

        if (tabBtns && tabBtns.length) {
            tabBtns.forEach(btn => {
                btn.addEventListener('click', (e) => {
                    this.switchTab(e.target.dataset.tab);
                });
            });
        }

        if (studentSearch) {
            studentSearch.addEventListener('input', () => {
                this.filterStudentTable();
            });
        }

        gradeSelect.addEventListener('change', () => {
            this.updateClassOptions();
            this.updateStudentOptions();
            this.filterStudentTable();
        });

        classSelect.addEventListener('change', () => {
            this.updateStudentOptions();
            this.filterStudentTable();
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

    handleStudentDetailKeydown(event) {
        if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
            return;
        }

        if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') {
            return;
        }

        const target = event.target;
        const tagName = target && target.tagName ? target.tagName.toLowerCase() : '';
        if (
            (target && target.isContentEditable) ||
            tagName === 'input' ||
            tagName === 'textarea' ||
            tagName === 'select'
        ) {
            return;
        }

        const studentsTab = document.getElementById('students-tab');
        const detailViewBtn = document.getElementById('detailViewBtn');
        const detailView = document.getElementById('detailView');
        const studentSelect = document.getElementById('studentSelect');

        const isStudentsTabActive = !!studentsTab && studentsTab.classList.contains('active');
        const isDetailViewActive = !!detailViewBtn && detailViewBtn.classList.contains('active');
        const isDetailViewVisible = !!detailView && detailView.style.display !== 'none';

        if (!isStudentsTabActive || !isDetailViewActive || !isDetailViewVisible || !studentSelect || !studentSelect.value) {
            return;
        }

        event.preventDefault();
        this.navigateStudentDetail(event.key === 'ArrowLeft' ? -1 : 1);
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
            await this.subjectGroupsReady;
            this.filesData.clear();
            
            for (const file of files) {
                const data = await this.readExcelFile(file);
                const fileData = this.parseFileData(data, file.name);
                this.filesData.set(file.name, fileData);
            }
            
            this.combineAllData();
            this.displayResults();
            this.hideLoading();

            // Enable export buttons after successful analysis
            const exportCsvBtn = document.getElementById('exportCsvBtn');
            const exportHtmlBtn = document.getElementById('exportHtmlBtn');
            if (exportCsvBtn) exportCsvBtn.disabled = false;
            if (exportHtmlBtn) exportHtmlBtn.disabled = false;
            
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

        // 모든 학생 데이터 통합 (같은 학년-반-번호 학생은 병합)
        const studentMap = new Map();

        this.filesData.forEach((fileData, fileName) => {
            fileData.students.forEach(student => {
                // 학생 고유 키: 학년-반-번호 (같은 학생 식별)
                const studentKey = `${fileData.grade}-${fileData.class}-${student.number}`;

                if (!studentMap.has(studentKey)) {
                    // 새로운 학생 생성
                    studentMap.set(studentKey, {
                        originalNumber: student.number,
                        originalName: student.name,
                        name: student.name,
                        displayName: `${fileData.grade}학년${fileData.class}반-${student.name}`,
                        grade: fileData.grade,
                        class: fileData.class,
                        fileNames: [fileName],
                        scores: {},
                        achievements: {},
                        grades: {},
                        ranks: {},
                        subjectTotals: {},
                        percentiles: {},
                        totalStudents: student.totalStudents,
                        hasGradeReportSource: fileData.format === 'grade-report',
                        hasXlsDataSource: fileData.format !== 'grade-report'
                    });
                } else {
                    // 기존 학생에 파일명 추가
                    studentMap.get(studentKey).fileNames.push(fileName);
                }

                const combinedStudent = studentMap.get(studentKey);
                combinedStudent.hasGradeReportSource = combinedStudent.hasGradeReportSource || fileData.format === 'grade-report';
                combinedStudent.hasXlsDataSource = combinedStudent.hasXlsDataSource || fileData.format !== 'grade-report';

                // 과목별 데이터 병합 (각 파일의 과목 데이터를 추가)
                Object.keys(student.scores || {}).forEach(subjectName => {
                    combinedStudent.scores[subjectName] = student.scores[subjectName];
                });
                Object.keys(student.achievements || {}).forEach(subjectName => {
                    combinedStudent.achievements[subjectName] = student.achievements[subjectName];
                });
                Object.keys(student.grades || {}).forEach(subjectName => {
                    combinedStudent.grades[subjectName] = student.grades[subjectName];
                });
                Object.keys(student.ranks || {}).forEach(subjectName => {
                    combinedStudent.ranks[subjectName] = student.ranks[subjectName];
                });
                Object.keys(student.subjectTotals || {}).forEach(subjectName => {
                    combinedStudent.subjectTotals[subjectName] = student.subjectTotals[subjectName];
                });

                // 수강자수 업데이트 (최대값 사용)
                if (student.totalStudents && (!combinedStudent.totalStudents || student.totalStudents > combinedStudent.totalStudents)) {
                    combinedStudent.totalStudents = student.totalStudents;
                }
            });
        });

        // Map을 배열로 변환하고 번호 할당
        let studentCounter = 1;
        studentMap.forEach((student, key) => {
            student.number = studentCounter++;
            // 병합 후 가중평균등급 재계산
            student.weightedAverageGrade = this.calculateWeightedAverageGrade(student, this.combinedData.subjects);
            this.combinedData.students.push(student);
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

            // 기본 분모: 실제 집계된 석차 보유자 수
            const totalStudents = studentsWithRanks.length;

            // 각 학생의 백분위 계산
            studentsWithRanks.forEach((item, index) => {
                const studentRank = item.rank;
                
                // 같은 석차의 학생들 찾기
                const sameRankStudents = studentsWithRanks.filter(s => s.rank === studentRank);
                const sameRankCount = sameRankStudents.length;
                
                // 해당 석차보다 나쁜 석차의 학생 수 (석차가 높은 학생들)
                const worseRankCount = studentsWithRanks.filter(s => s.rank > studentRank).length;
                
                // 분모 선택: 과목별 수강자수(subjectTotals)가 있으면 그 값을 우선 사용
                const subjTotal = item.student.subjectTotals && item.student.subjectTotals[subject.name]
                    ? item.student.subjectTotals[subject.name]
                    : totalStudents;
                // 백분위 계산(동점 보정): (전체 - 석차 + 0.5) / 전체 * 100
                const raw = ((subjTotal - studentRank + 0.5) / Math.max(1, subjTotal)) * 100;
                const percentile = raw;
                
                // 0~100 범위로 제한하고 내림 처리하여 경계 상향 편향 방지
                const finalPercentile = Math.max(0, Math.min(100, Math.floor(percentile)));
                
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
            if (student.weightedAverage9Grade === null || student.weightedAverage9Grade === undefined) {
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
        const format = this.detectFileFormat(data);
        if (format === 'grade-report') {
            console.log(`[${fileName}] 인쇄용 성적표 양식 감지`);
            return this.parseGradeReport(data, fileName);
        }

        const fileData = {
            fileName: fileName,
            data: data,
            subjects: [],
            students: [],
            grade: 1,
            class: 1,
            format: 'xls-data'
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

    detectFileFormat(data) {
        if (!data || data.length < 5) return 'xls-data';

        const row4 = data[3];
        if (!row4 || row4.length < 4) return 'xls-data';

        const knownHeaders = [
            '번호', '성명', '학년', '학기', '교과',
            '과목명', '과목', '학점', '단위수',
            '석차등급', '수강자수', '성취도', '원점수'
        ];
        let matchCount = 0;

        for (let c = 0; c < Math.min(row4.length, 20); c++) {
            const cell = String(row4[c] || '').replace(/\s+/g, '').trim();
            if (knownHeaders.some(header => cell.includes(header))) {
                matchCount++;
            }
        }

        if (matchCount >= 4) return 'grade-report';

        for (let c = 3; c < row4.length; c++) {
            const cell = String(row4[c] || '').trim();
            if (/^.+\(\d+\)$/.test(cell)) {
                return 'xls-data';
            }
        }

        return 'xls-data';
    }

    parseGradeReport(data, fileName) {
        const fileData = {
            fileName: fileName,
            data: data,
            subjects: [],
            students: [],
            grade: 1,
            class: 1,
            format: 'grade-report'
        };

        if (data[2] && data[2][0]) {
            const info = String(data[2][0]);
            const gradeMatches = info.match(/(\d+)\s*학년/g);
            if (gradeMatches) {
                const lastMatch = gradeMatches[gradeMatches.length - 1];
                const gradeMatch = lastMatch.match(/(\d+)/);
                if (gradeMatch) {
                    const parsedGrade = parseInt(gradeMatch[1], 10);
                    if (parsedGrade < 10) fileData.grade = parsedGrade;
                }
            }

            const classMatch = info.match(/(\d+)\s*반/);
            if (classMatch) {
                fileData.class = parseInt(classMatch[1], 10);
            }

            console.log(`[인쇄용 양식] ${fileData.grade}학년 ${fileData.class}반 감지`);
        }

        const headerRow = data[3] || [];
        const colMap = this._buildGradeReportColumnMap(headerRow);
        console.log('[인쇄용 양식] 열 매핑:', JSON.stringify(colMap));

        let curNumber = null;
        let curName = null;
        let curSchoolYear = null;
        let curSemester = null;
        let is예체능 = false;
        let is진로선택 = false;

        const subjectMap = new Map();
        const studentMap = new Map();
        const subjectOrder = [];

        for (let i = 4; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;

            const cellA = String(row[0] || '').trim();

            if (cellA.includes('체육') && (cellA.includes('예술') || cellA.includes('과학탐구실험'))) {
                is예체능 = true;
                continue;
            }
            if (cellA.includes('진로') && cellA.includes('선택')) {
                is진로선택 = true;
                continue;
            }
            if (cellA.startsWith('<') && !cellA.includes('체육') && !cellA.includes('진로')) {
                // 섹션 구분선은 그대로 건너뛴다.
                continue;
            }

            if (this._isGradeReportHeaderRow(row, colMap)) continue;

            const subjectName = this._grVal(row, colMap, 'subjectName');
            const creditsRaw = this._grVal(row, colMap, 'credits');
            if (!subjectName || String(subjectName).trim() === '') continue;

            const credits = parseFloat(creditsRaw);
            if (isNaN(credits)) continue;

            const numVal = this._grVal(row, colMap, 'number');
            if (numVal !== undefined && numVal !== null && numVal !== '') {
                const parsedNumber = parseInt(numVal, 10);
                if (!isNaN(parsedNumber)) {
                    curNumber = parsedNumber;
                    const nameVal = this._grVal(row, colMap, 'name');
                    if (nameVal && String(nameVal).trim() !== '') {
                        curName = String(nameVal).trim();
                    }
                }
            }

            const yearVal = this._grVal(row, colMap, 'schoolYear');
            if (yearVal !== undefined && yearVal !== null && yearVal !== '') {
                const parsedYear = parseInt(yearVal, 10);
                if (!isNaN(parsedYear)) curSchoolYear = parsedYear;
            }

            const semVal = this._grVal(row, colMap, 'semester');
            if (semVal !== undefined && semVal !== null && semVal !== '') {
                const parsedSemester = parseInt(semVal, 10);
                if (!isNaN(parsedSemester)) curSemester = parsedSemester;
            }

            if (curNumber === null) continue;

            const subName = String(subjectName).trim();
            const subjectGroup = String(this._grVal(row, colMap, 'subjectGroup') || '').trim();

            let rawScore = null;
            let subjectAvg = 0;
            const rawScoreCell = this._grVal(row, colMap, 'rawScore');
            const avgCell = this._grVal(row, colMap, 'subjectAvg');

            if (rawScoreCell !== undefined && rawScoreCell !== null && rawScoreCell !== '') {
                const rawStr = String(rawScoreCell).trim();
                if (rawStr.includes('/')) {
                    const parts = rawStr.split('/');
                    const parsedRawScore = parseFloat(parts[0]);
                    rawScore = isNaN(parsedRawScore) ? null : parsedRawScore;
                    if (parts[1]) {
                        subjectAvg = parseFloat(parts[1].split('(')[0]) || 0;
                    }
                } else {
                    const parsedRawScore = parseFloat(rawStr);
                    rawScore = isNaN(parsedRawScore) ? null : parsedRawScore;
                }
            }

            if (avgCell !== undefined && avgCell !== null && avgCell !== '' && subjectAvg === 0) {
                subjectAvg = parseFloat(avgCell) || 0;
            }

            let achievement = '';
            if (is예체능 && colMap.achievement !== undefined) {
                const achVal = this._grVal(row, colMap, 'achievement');
                achievement = this._normalizeAchievementValue(achVal);
                if (!achievement && colMap.achievement > 0) {
                    achievement = this._normalizeAchievementValue(row[colMap.achievement - 1]);
                }
            } else {
                achievement = this._normalizeAchievementValue(this._grVal(row, colMap, 'achievement'));
            }

            let gradeRank = NaN;
            if (!is예체능) {
                const gradeRankRaw = this._grVal(row, colMap, 'gradeRank');
                if (gradeRankRaw !== undefined && gradeRankRaw !== null && gradeRankRaw !== '') {
                    const gradeMatch = String(gradeRankRaw).trim().match(/\d+/);
                    if (gradeMatch) gradeRank = parseInt(gradeMatch[0], 10);
                }
            }

            let totalStudents = NaN;
            if (!is예체능) {
                const totalRaw = this._grVal(row, colMap, 'totalStudents');
                if (totalRaw !== undefined && totalRaw !== null && totalRaw !== '') {
                    const totalMatch = String(totalRaw).trim().match(/\d+/);
                    if (totalMatch) totalStudents = parseInt(totalMatch[0], 10);
                }
            }

            const distRaw = this._grVal(row, colMap, 'achievementDist');

            if (!subjectMap.has(subName)) {
                subjectMap.set(subName, {
                    name: subName,
                    credits: credits,
                    averages: [],
                    rawDistributions: [],
                    group: subjectGroup,
                    isCareerTrack: is진로선택,
                    schoolYear: curSchoolYear,
                    semester: curSemester
                });
                subjectOrder.push(subName);
            }

            const subjectInfo = subjectMap.get(subName);
            if (subjectAvg > 0) subjectInfo.averages.push(subjectAvg);
            if (distRaw && String(distRaw).trim() !== '') {
                subjectInfo.rawDistributions.push(String(distRaw).trim());
            }

            if (!studentMap.has(curNumber)) {
                studentMap.set(curNumber, {
                    number: curNumber,
                    name: curName || `학생${curNumber}`,
                    scores: {},
                    achievements: {},
                    grades: {},
                    ranks: {},
                    subjectTotals: {},
                    percentiles: {},
                    totalStudents: null,
                    sourceFormat: fileData.format,
                    hasGradeReportSource: true,
                    hasXlsDataSource: false
                });
            }

            const student = studentMap.get(curNumber);
            if (curName && curName !== `학생${curNumber}`) {
                student.name = curName;
            }

            if (rawScore !== null) {
                student.scores[subName] = rawScore;
            }
            if (achievement) student.achievements[subName] = achievement;

            if (!is예체능) {
                if (!isNaN(gradeRank)) student.grades[subName] = gradeRank;
                if (!isNaN(totalStudents)) {
                    student.subjectTotals[subName] = totalStudents;
                    if (!student.totalStudents || totalStudents > student.totalStudents) {
                        student.totalStudents = totalStudents;
                    }
                }
            }
        }

        subjectOrder.forEach((subName, idx) => {
            const info = subjectMap.get(subName);
            const subject = {
                name: info.name,
                credits: info.credits,
                columnIndex: idx,
                average: info.averages.length > 0
                    ? info.averages.reduce((sum, value) => sum + value, 0) / info.averages.length
                    : 0,
                scores: []
            };

            if (info.rawDistributions.length > 0) {
                subject.distribution = this._parseAchievementDistString(info.rawDistributions[0]);
            }

            fileData.subjects.push(subject);
        });

        studentMap.forEach(student => {
            student.weightedAverageGrade = this.calculateWeightedAverageGrade(student, fileData.subjects);
            student.weightedAverage9Grade = this.calculateWeightedAverage9Grade(student, fileData.subjects);
            fileData.students.push(student);
        });

        console.log(`[인쇄용 양식] 과목 ${fileData.subjects.length}개, 학생 ${fileData.students.length}명 파싱 완료`);
        return fileData;
    }

    _buildGradeReportColumnMap(headerRow) {
        const colMap = {};
        const nameMap = [
            { keys: ['번호'], field: 'number' },
            { keys: ['성명', '이름'], field: 'name' },
            { keys: ['학년'], field: 'schoolYear' },
            { keys: ['학기'], field: 'semester' },
            { keys: ['교과'], field: 'subjectGroup' },
            { keys: ['과목명', '과목'], field: 'subjectName' },
            { keys: ['학점', '단위수', '단위'], field: 'credits' },
            { keys: ['원점수'], field: 'rawScore' },
            { keys: ['과목평균'], field: 'subjectAvg' },
            { keys: ['석차등급'], field: 'gradeRank' },
            { keys: ['수강자수'], field: 'totalStudents' },
            { keys: ['성취도별분포비율', '성취도별 분포비율', '분포비율'], field: 'achievementDist' }
        ];

        for (let c = 0; c < headerRow.length; c++) {
            const raw = String(headerRow[c] || '').replace(/\s+/g, '').trim();
            if (!raw) continue;

            for (const mapping of nameMap) {
                if (colMap[mapping.field] !== undefined) continue;
                for (const key of mapping.keys) {
                    if (raw === key.replace(/\s+/g, '')) {
                        colMap[mapping.field] = c;
                        break;
                    }
                }
            }
        }

        for (let c = 0; c < headerRow.length; c++) {
            const raw = String(headerRow[c] || '').replace(/\s+/g, '').trim();
            if (!raw) continue;

            for (const mapping of nameMap) {
                if (colMap[mapping.field] !== undefined) continue;
                for (const key of mapping.keys) {
                    if (raw.includes(key.replace(/\s+/g, ''))) {
                        colMap[mapping.field] = c;
                        break;
                    }
                }
            }
        }

        if (colMap.achievement === undefined) {
            for (let c = 0; c < headerRow.length; c++) {
                const raw = String(headerRow[c] || '').replace(/\s+/g, '').trim();
                if (raw.includes('성취도') && !raw.includes('분포') && !raw.includes('비율')) {
                    colMap.achievement = c;
                    break;
                }
            }
        }

        if (colMap.subjectAvg === undefined) {
            for (let c = 0; c < headerRow.length; c++) {
                const raw = String(headerRow[c] || '').replace(/\s+/g, '').trim();
                if (raw === '평균' && c !== colMap.rawScore) {
                    colMap.subjectAvg = c;
                    break;
                }
            }
        }

        if (colMap.rawScore === undefined && colMap.credits !== undefined) {
            colMap.rawScore = colMap.credits + 1;
        }

        return colMap;
    }

    _isGradeReportHeaderRow(row, colMap) {
        const cellA = String(row[0] || '').trim();
        if (cellA === '번호') return true;

        if (colMap.subjectName !== undefined) {
            const subjectName = String(row[colMap.subjectName] || '').trim();
            if (subjectName === '과목명' || subjectName === '과목') return true;
        }

        if (colMap.credits !== undefined) {
            const credits = String(row[colMap.credits] || '').trim();
            if (credits === '학점' || credits === '단위수') return true;
        }

        return false;
    }

    _grVal(row, colMap, field) {
        if (colMap[field] === undefined) return undefined;
        return row[colMap[field]];
    }

    _parseAchievementDistString(str) {
        const distribution = {};
        if (!str) return distribution;

        const matches = str.match(/[ABCDE]\s*\(\s*\d+\.?\d*\s*\)/g);
        if (matches) {
            matches.forEach(match => {
                const parsed = match.match(/([ABCDE])\s*\(\s*(\d+\.?\d*)\s*\)/);
                if (parsed) distribution[parsed[1]] = parseFloat(parsed[2]);
            });
        }

        return distribution;
    }

    _normalizeAchievementValue(value) {
        if (value === undefined || value === null) return '';

        const normalized = String(value).trim();
        if (!normalized || normalized.includes('전입')) return '';

        const match = normalized.match(/^[ABCDE]/);
        return match ? match[0] : '';
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
                subjectTotals: {},
                percentiles: {},
                sourceFormat: fileData.format,
                hasGradeReportSource: fileData.format === 'grade-report',
                hasXlsDataSource: fileData.format !== 'grade-report',
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
                
                // 석차등급 (문자 혼입 시 숫자만 추출)
                if (gradeRow && gradeRow[colIndex] !== undefined && gradeRow[colIndex] !== null) {
                    const gradeText = String(gradeRow[colIndex]).trim();
                    const gm = gradeText.match(/\d+/);
                    if (gm) {
                        student.grades[subject.name] = parseInt(gm[0], 10);
                    }
                }

                // 석차 (동석차 표기 포함 대비: 숫자만 추출)
                if (rankRow && rankRow[colIndex] !== undefined && rankRow[colIndex] !== null) {
                    const rankText = String(rankRow[colIndex]).trim();
                    const rm = rankText.match(/\d+/);
                    if (rm) {
                        student.ranks[subject.name] = parseInt(rm[0], 10);
                    }
                }

                // 수강자수 (과목별로 저장) 숫자만 추출
                if (totalRow && totalRow[colIndex] !== undefined && totalRow[colIndex] !== null) {
                    const totalText = String(totalRow[colIndex]).trim();
                    const tm = totalText.match(/\d+/);
                    if (tm) {
                        const total = parseInt(tm[0], 10);
                        student.subjectTotals[subject.name] = total;
                        // 기존 totalStudents는 호환을 위해 첫 과목에서만 설정 (전체 학생 수 표시용)
                        if (!student.totalStudents) {
                            student.totalStudents = total;
                        }
                    }
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
    
        // 1차: 석차 기반 백분위 (기존 로직)
        if (student.percentiles && student.ranks) {
            subjects.forEach(subject => {
                const percentile = student.percentiles[subject.name];
                const rank = student.ranks[subject.name];
                if (percentile !== undefined && percentile !== null 
                    && rank !== undefined && rank !== null && !isNaN(rank)) {
                    totalPercentilePoints += percentile * subject.credits;
                    totalCredits += subject.credits;
                }
            });
        }
    
        if (totalCredits > 0) {
            return totalPercentilePoints / totalCredits;
        }
    
        // 2차 폴백: 석차 없이 등급만 있는 경우 (grade-report 양식)
        // 등급 + 수강자수 기반 백분위 추정
        if (student.grades) {
            subjects.forEach(subject => {
                const grade = student.grades[subject.name];
                if (grade !== undefined && grade !== null && !isNaN(grade)) {
                    const subjectTotal = (student.subjectTotals && student.subjectTotals[subject.name])
                        ? student.subjectTotals[subject.name]
                        : null;
                    const estimatedPercentile = this.estimatePercentileFromGrade(grade, subjectTotal);
                    if (estimatedPercentile !== null) {
                        totalPercentilePoints += estimatedPercentile * subject.credits;
                        totalCredits += subject.credits;
                    }
                }
            });
        }
    
        return totalCredits > 0 ? totalPercentilePoints / totalCredits : null;
    }

    /**
     * 5등급(석차등급)과 수강자수로부터 백분위를 추정한다.
     * 
     * 2022 개정 교육과정 5등급 성취평가제 기준 누적비율:
     *   A: 상위 ~0%  ~ 누적별도(성취수준별 인원비율은 학교마다 다름)
     * 
     * 석차등급(1~5)의 경우, 9등급 누적비율 매핑 기준의 구간 중앙값을 사용한다.
     * 수강자수가 있으면 해당 등급의 누적비율 중앙에서 좀 더 정밀하게 추정한다.
     *
     * @param {number} grade - 석차등급 (1~5 또는 1~9)
     * @param {number|null} totalStudents - 수강자수 (없으면 null)
     * @returns {number|null} 추정 백분위 (0~100)
     */
    estimatePercentileFromGrade(grade, totalStudents) {
        if (grade === null || grade === undefined || isNaN(grade)) {
            return null;
        }

        // 5등급 체계 (석차등급 1~5)
        // 각 등급의 백분위 구간 [하한, 상한] 및 중앙값
        const fiveGradePercentileMap = {
            1: { lower: 90, upper: 100, mid: 96 },  // 상위 ~10% → 백분위 90~100
            2: { lower: 70, upper: 90,  mid: 82 },   // 상위 ~30% → 백분위 70~90
            3: { lower: 40, upper: 70,  mid: 58 },   // 상위 ~60% → 백분위 40~70
            4: { lower: 10, upper: 40,  mid: 28 },   // 상위 ~90% → 백분위 10~40
            5: { lower: 0,  upper: 10,  mid: 5 }     // 하위 ~10% → 백분위 0~10
        };

        const rounded = Math.round(grade);
        const mapping = fiveGradePercentileMap[rounded];
    
        if (!mapping) {
            // 9등급 체계인 경우 대비
            return this.convertPercentileTo9Grade 
                ? null  // 9등급은 별도 처리
                : null;
        }

        return mapping.mid;
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

    getBusanGradeAverageToNineGradeTable() {
        return [
            { grade5: 1.08, grade9: 1.59 },
            { grade5: 1.16, grade9: 1.78 },
            { grade5: 1.24, grade9: 1.98 },
            { grade5: 1.33, grade9: 2.14 },
            { grade5: 1.42, grade9: 2.32 },
            { grade5: 1.50, grade9: 2.45 },
            { grade5: 1.66, grade9: 2.72 },
            { grade5: 1.83, grade9: 3.03 },
            { grade5: 2.00, grade9: 3.35 },
            { grade5: 2.16, grade9: 3.60 },
            { grade5: 2.33, grade9: 3.91 },
            { grade5: 2.50, grade9: 4.20 },
            { grade5: 2.66, grade9: 4.46 },
            { grade5: 2.83, grade9: 4.73 },
            { grade5: 3.00, grade9: 5.03 },
            { grade5: 3.16, grade9: 5.28 },
            { grade5: 3.33, grade9: 5.58 },
            { grade5: 3.50, grade9: 5.86 },
            { grade5: 3.66, grade9: 6.08 },
            { grade5: 3.83, grade9: 6.37 },
            { grade5: 4.00, grade9: 6.67 },
            { grade5: 4.16, grade9: 6.93 },
            { grade5: 4.33, grade9: 7.20 },
            { grade5: 4.50, grade9: 7.48 },
            { grade5: 4.66, grade9: 7.71 },
            { grade5: 4.83, grade9: 8.00 },
            { grade5: 5.00, grade9: 9.00 }
        ];
    }

    estimateNineGradeAverageFromFiveGradeAverage(gradeAverage) {
        if (gradeAverage === null || gradeAverage === undefined || isNaN(gradeAverage)) {
            return null;
        }

        const table = this.getBusanGradeAverageToNineGradeTable();
        if (table.length === 0) return null;

        if (gradeAverage <= table[0].grade5) {
            return table[0].grade9;
        }

        const lastPoint = table[table.length - 1];
        if (gradeAverage >= lastPoint.grade5) {
            return lastPoint.grade9;
        }

        for (let i = 1; i < table.length; i++) {
            const prev = table[i - 1];
            const next = table[i];

            if (gradeAverage === next.grade5) {
                return next.grade9;
            }

            if (gradeAverage < next.grade5) {
                const ratio = (gradeAverage - prev.grade5) / (next.grade5 - prev.grade5);
                return prev.grade9 + ((next.grade9 - prev.grade9) * ratio);
            }
        }

        return lastPoint.grade9;
    }

    // (제거됨) 5등급 기반 9등급 하한 강제 로직은 오류 탐지 가시성을 해치므로 사용하지 않음

    calculateExactWeightedAverage9Grade(student, subjects) {
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

    // 9등급 가중평균 계산
    calculateWeightedAverage9Grade(student, subjects) {
        const exactWeightedAverage9Grade = this.calculateExactWeightedAverage9Grade(student, subjects);
        if (exactWeightedAverage9Grade !== null) {
            return exactWeightedAverage9Grade;
        }

        const isGradeReportSource = student &&
            (student.hasGradeReportSource || student.sourceFormat === 'grade-report');

        if (!isGradeReportSource) {
            return null;
        }

        return this.estimateNineGradeAverageFromFiveGradeAverage(student.weightedAverageGrade);
    }

    usesBusanNineGradeReference(student, subjects) {
        if (!student || student.weightedAverage9Grade === null || student.weightedAverage9Grade === undefined) {
            return false;
        }

        const isGradeReportSource = student.hasGradeReportSource || student.sourceFormat === 'grade-report';
        if (!isGradeReportSource) {
            return false;
        }

        return this.calculateExactWeightedAverage9Grade(student, subjects) === null;
    }


    displayResults() {
        document.getElementById('results').style.display = 'block';
        this.displaySubjectAverages();
        this.displayGradeAnalysis();
        this.displayStudentAnalysis();
        if (document.querySelector('[data-tab="subjects"]') && document.getElementById('subjects-tab')) {
            this.switchTab('subjects');
        }
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
                <div class="credits">2026 강원진학센터 입시분석팀 남궁연(강원 설악고등학교)</div>
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
                "2026 강원진학센터 입시분석팀 남궁연(강원 설악고등학교)\\n" +
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
                const readme = "배포용 성적 분석 뷰어\\n========================\\n\\n사용법:\\n1. 모든 파일을 같은 폴더에 저장하세요\\n2. index.html 파일을 웹브라우저에서 열어주세요\\n\\n2026 강원진학센터 입시분석팀 남궁연(강원 설악고등학교)\\n링크: https://namgungyeon.tistory.com/133";
                this.downloadFile(readme, "README.txt", "text/plain");
            }, 1500);
            
            alert(`배포용 파일들을 다운로드하고 있습니다.\\n\\n모든 파일을 같은 폴더에 저장한 후\\nindex.html 파일을 열어서 사용하세요.`);
        }
    }

    // Export HTML that references external style.css and script.js (paired files)
    async exportAsPairedHtml() {
        if (!this.combinedData) {
            this.showError('먼저 파일을 분석하세요.');
            return;
        }
        // Helper fetch
        const safeFetchText = async (url) => {
            try {
                const res = await fetch(url, { cache: 'no-cache' });
                if (!res.ok) throw new Error('HTTP ' + res.status);
                return await res.text();
            } catch (_) { return ''; }
        };

        // 1) index.html 생성 (원본 파일 선호, 실패 시 현재 문서 기반) + PRELOADED_DATA 주입
        const parser = new DOMParser();
        let indexSrc = await (async () => {
            try {
                const res = await fetch('index.html', { cache: 'no-cache' });
                if (res && res.ok) return await res.text();
            } catch (_) {}
            return document.documentElement.outerHTML;
        })();
        const doc = parser.parseFromString(indexSrc, 'text/html');
        const preload = doc.createElement('script');
        preload.textContent = `window.APP_BUILD_UTC = new Date().toISOString();\nwindow.PRELOADED_DATA = ${JSON.stringify(this.combinedData)};`;
        const appScript = doc.querySelector('script[src="script.js"]');
        if (appScript) appScript.before(preload); else { doc.body.appendChild(preload); const s = doc.createElement('script'); s.src = 'script.js'; doc.body.appendChild(s); }
        const indexOut = '<!DOCTYPE html>' + doc.documentElement.outerHTML;

        // 2) 현재 style.css, script.js 내용 확보 (정확히 동일 파일을 사용 - 실패 시 에러 표시)
        let cssText = await safeFetchText('style.css');
        let jsText = await safeFetchText('script.js');
        // fetch 실패 시, 사용자가 로컬 파일을 직접 선택해서 복사할 수 있도록 안내
        if ((!cssText || !jsText) && window.showOpenFilePicker) {
            try {
                if (!cssText) {
                    const [cssHandle] = await window.showOpenFilePicker({
                        multiple: false,
                        types: [{ description: 'CSS', accept: { 'text/css': ['.css'] } }]
                    });
                    const cssFile = await cssHandle.getFile();
                    cssText = await cssFile.text();
                }
            } catch (e) { /* 사용자가 취소한 경우 등은 무시 */ }
            try {
                if (!jsText) {
                    const [jsHandle] = await window.showOpenFilePicker({
                        multiple: false,
                        types: [{ description: 'JavaScript', accept: { 'application/javascript': ['.js'] } }]
                    });
                    const jsFile = await jsHandle.getFile();
                    jsText = await jsFile.text();
                }
            } catch (e) { /* 무시 */ }
        }
        if (!cssText || !jsText) {
            console.warn('원본 style.css/script.js를 일부 가져오지 못했습니다. ZIP에는 빈 파일이 포함될 수 있습니다.');
        }

        // 3) 항상 ZIP으로 같은 폴더 평면 구조로 다운로드
        const zip = new JSZip();
        zip.file('index.html', indexOut);
        zip.file('style.css', cssText || '/* style */');
        zip.file('script.js', jsText || '/* script */');
        const blob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        a.download = `analysis_${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}.zip`;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
        return;
    }

    async generateExactSnapshotHtmlTemplate() {
        // 차트가 모두 그려지도록 보장 (애니메이션 없이 최신 상태로 업데이트)
        await this.ensureChartsRendered();
        // 렌더 안정화 대기(레이아웃/폰트/애니메이션 마무리)
        await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
        await new Promise(r => setTimeout(r, 200));

        const cssContent = await this.getStyleCSS();
        const container = document.querySelector('.container');
        if (!container) throw new Error('내보낼 컨테이너를 찾을 수 없습니다.');

        const containerClone = container.cloneNode(true);
        const origCanvases = container.querySelectorAll('canvas');
        const cloneCanvases = containerClone.querySelectorAll('canvas');

        for (let i = 0; i < cloneCanvases.length; i++) {
            const srcCanvas = origCanvases[i];
            const dstCanvas = cloneCanvases[i];
            if (srcCanvas && dstCanvas && srcCanvas.toDataURL) {
                try {
                    const img = document.createElement('img');
                    img.src = srcCanvas.toDataURL('image/png');
                    const rect = srcCanvas.getBoundingClientRect();
                    img.style.width = Math.max(1, Math.round(rect.width)) + 'px';
                    img.style.height = Math.max(1, Math.round(rect.height)) + 'px';
                    img.className = dstCanvas.className || '';
                    if (dstCanvas.id) img.id = dstCanvas.id;
                    img.alt = dstCanvas.getAttribute('aria-label') || 'chart-image';
                    dstCanvas.replaceWith(img);
                } catch (_) {
                    // 실패 시 캔버스 그대로 둔다.
                }
            }
        }

        const title = document.title || '(2022개정) 고등학교 내신 분석 프로그램 Lite';
        return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <style>
${cssContent}
  </style>
</head>
<body>
${containerClone.outerHTML}
</body>
</html>`;
    }

    // 현재 화면 상태 그대로(차트 포함) 정적인 HTML로 저장
    async exportAsExactSnapshotHtml() {
        if (!this.combinedData) {
            this.showError('먼저 파일을 분석하세요.');
            return;
        }

        try {
            const html = await this.generateExactSnapshotHtmlTemplate();

            // 다운로드 (BOM 포함: 한글 표시 안전)
            const BOM = '\uFEFF';
            const blob = new Blob([BOM + html], { type: 'text/html;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            const now = new Date();
            const pad = (n) => String(n).padStart(2, '0');
            const filename = `학생성적분석_스냅샷_${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}.html`;
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 0);

        } catch (err) {
            console.error('스냅샷 HTML 생성 오류:', err);
            this.showError('스냅샷 HTML 생성 중 오류가 발생했습니다: ' + (err && err.message ? err.message : String(err)));
        }
    }

    async ensureChartsRendered() {
        try {
            if (this.scatterChart && typeof this.scatterChart.update === 'function') {
                this.scatterChart.update('none');
            }
        } catch (_) {}
        try {
            if (this.barChart && typeof this.barChart.update === 'function') {
                this.barChart.update('none');
            }
        } catch (_) {}
        try {
            if (this.studentPercentileChart && typeof this.studentPercentileChart.update === 'function') {
                this.studentPercentileChart.update('none');
            }
        } catch (_) {}
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
        // CSS 로드가 실패했을 때 사용할 기본 스타일 (style.css와 동기화됨)
        return `/* ========================================
   Modern Clean Theme
   ======================================== */

:root {
    --primary: #5F4A8B;
    --primary-light: #7B62A8;
    --primary-dark: #483670;
    --primary-bg: rgba(95, 74, 139, 0.08);

    --accent: #7B62A8;
    --accent-light: #9578C4;
    --accent-muted: #E8E0F4;
    --accent-bg: rgba(123, 98, 168, 0.08);

    --neutral-50: #FCFCFD;
    --neutral-100: #F8F7FA;
    --neutral-200: #EEEDF2;
    --neutral-300: #DDDBE5;
    --neutral-400: #B8B4C4;
    --neutral-500: #7E7891;
    --neutral-600: #585269;
    --neutral-700: #36314A;
    --neutral-800: #1A1626;

    --success: #15803D;
    --success-light: #16A34A;
    --success-bg: rgba(21, 128, 61, 0.1);

    --info: #2563EB;
    --info-light: #3B82F6;
    --info-bg: rgba(37, 99, 235, 0.1);

    --warning: #D97706;
    --warning-light: #F59E0B;
    --warning-bg: rgba(217, 119, 6, 0.12);

    --bg-body: radial-gradient(circle at top, #FFFFFF 0%, #F6F5F9 42%, #EEEDF2 100%);
    --bg-card: #FFFFFF;
    --bg-card-hover: #FFFFFF;
    --bg-section: linear-gradient(180deg, #FFFFFF 0%, #F8F7FA 100%);

    --text-primary: #1A1626;
    --text-secondary: #585269;
    --text-muted: #7E7891;
    --text-inverse: #FFFFFF;

    --border-light: rgba(95, 74, 139, 0.08);
    --border-medium: rgba(95, 74, 139, 0.14);
    --border-accent: rgba(95, 74, 139, 0.18);

    --shadow-sm: 0 1px 2px rgba(95, 74, 139, 0.06);
    --shadow-md: 0 8px 24px rgba(95, 74, 139, 0.08);
    --shadow-lg: 0 16px 36px rgba(95, 74, 139, 0.10);
    --shadow-xl: 0 24px 64px rgba(95, 74, 139, 0.12);

    --radius-sm: 10px;
    --radius-md: 14px;
    --radius-lg: 20px;
    --radius-xl: 28px;

    --grade-1: #15803D;
    --grade-2: #5F4A8B;
    --grade-3: #0369A1;
    --grade-4: #D97706;
    --grade-5: #7E7891;
}

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: 'Pretendard Variable', 'Pretendard', 'SUIT Variable', 'Noto Sans KR', sans-serif;
    background: var(--bg-body);
    min-height: 100vh;
    padding: 18px;
    position: relative;
    overflow-x: hidden;
    color: var(--text-primary);
}

.container {
    max-width: 1320px;
    margin: 0 auto;
    background: var(--bg-card);
    border-radius: var(--radius-xl);
    box-shadow: var(--shadow-xl);
    overflow: hidden;
    position: relative;
    border: 1px solid var(--border-light);
}

header {
    background: linear-gradient(135deg, #5F4A8B 0%, #483670 50%, #36314A 100%);
    color: var(--text-inverse);
    padding: 40px 36px 36px;
    text-align: center;
    border-bottom: none;
    position: relative;
    overflow: hidden;
}

header::before {
    content: '✦';
    position: absolute;
    top: 16px;
    right: 24px;
    font-size: 1.5rem;
    opacity: 0.3;
    color: #FFFFFF;
}

header h1 {
    font-size: 2rem;
    margin-bottom: 6px;
    font-weight: 600;
    letter-spacing: -0.04em;
    position: relative;
    color: #FFFFFF;
}

.header-subtitle {
    font-size: 1rem;
    color: rgba(255, 255, 255, 0.85);
    max-width: 720px;
    line-height: 1.6;
    margin: 0 auto;
}

.badge-lite {
    display: inline-block;
    margin-left: 10px;
    padding: 4px 10px;
    font-size: 0.74rem;
    font-weight: 600;
    color: #FFFFFF;
    background: rgba(255, 255, 255, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.4);
    border-radius: 999px;
    letter-spacing: 0.08em;
    vertical-align: middle;
    text-transform: uppercase;
}

.upload-section {
    padding: 32px 36px;
    text-align: center;
    border-bottom: 1px solid var(--neutral-200);
    position: relative;
    background: linear-gradient(180deg, var(--neutral-50) 0%, var(--bg-card) 100%);
}

.container.post-analysis .upload-guide,
.container.post-analysis .section-divider,
.container.post-analysis .file-input-wrapper,
.container.post-analysis #fileList,
.container.post-analysis #analyzeBtn {
    display: none !important;
}

.file-input-wrapper {
    margin-bottom: 30px;
}

.file-input-wrapper input[type="file"] {
    display: none;
}

.file-input-label {
    display: flex;
    align-items: center;
    justify-content: center;
    width: min(100%, 720px);
    min-height: 88px;
    margin: 0 auto;
    padding: 18px 28px;
    background: var(--bg-card);
    border: 1.5px dashed var(--neutral-300);
    border-radius: var(--radius-lg);
    cursor: pointer;
    transition: all 0.25s ease;
    font-size: 1rem;
    font-weight: 500;
    color: var(--text-primary);
}

.upload-hint {
    margin-top: 12px;
    font-size: 0.9rem;
    color: var(--text-muted);
}

.file-input-label:hover {
    background: var(--neutral-50);
    border-color: var(--primary);
    color: var(--primary-dark);
}

.file-input-label.dragover {
    background: var(--primary-bg);
    border-color: var(--primary);
    color: var(--primary);
    box-shadow: 0 0 0 4px var(--primary-bg) inset;
}

/* 업로드 섹션 전체 드래그오버 강조 및 오버레이 안내 */
.upload-section.dragover {
    border: 1.5px dashed var(--primary);
    border-radius: var(--radius-lg);
    background: var(--primary-bg);
}
.upload-section.dragover::after {
    content: '여기에 파일을 드롭하세요';
    position: absolute;
    left: 50%;
    top: 50%;
    transform: translate(-50%, -50%);
    color: var(--primary);
    font-weight: 600;
    font-size: 1rem;
    padding: 12px 18px;
    background: var(--bg-card);
    border: 1px solid var(--primary-light);
    border-radius: var(--radius-sm);
    pointer-events: none;
    box-shadow: var(--shadow-lg);
}

.analyze-btn {
    background: var(--primary);
    color: var(--text-inverse);
    border: 1px solid transparent;
    padding: 12px 22px;
    border-radius: 999px;
    font-size: 0.95rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.25s ease;
    box-shadow: none;
}

.analyze-btn:hover:not(:disabled) {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
    background: var(--primary-dark);
}

.analyze-btn:active:not(:disabled) {
    transform: translateY(0);
}

.analyze-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
    background: var(--neutral-400);
    box-shadow: none;
}

.action-buttons {
    display: flex;
    justify-content: center;
    gap: 10px;
    flex-wrap: wrap;
}

.secondary-btn {
    background: var(--bg-card);
    color: var(--text-primary);
    border-color: var(--neutral-300);
}

.secondary-btn:hover:not(:disabled) {
    background: var(--neutral-100);
    color: var(--text-primary);
    box-shadow: var(--shadow-sm);
}

.export-btn {
    background: var(--bg-card);
    color: var(--primary);
    border: 2px solid var(--primary);
    padding: 13px 32px;
    border-radius: var(--radius-lg);
    font-size: 1rem;
    cursor: pointer;
    transition: all 0.25s ease;
}

.export-btn:hover:not(:disabled) {
    background: var(--primary);
    color: var(--text-inverse);
    transform: translateY(-2px);
    box-shadow: var(--shadow-lg);
}

.export-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}

.upload-guide {
    background: var(--bg-card);
    border: 1px solid var(--neutral-200);
    padding: 22px 24px;
    margin: 0 auto 18px;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow-sm);
    max-width: 920px;
    text-align: left;
}

.upload-guide p {
    margin: 0;
    color: var(--text-secondary);
    line-height: 1.6;
}

.guide-title {
    color: var(--text-primary);
    margin-bottom: 8px !important;
    font-size: 1rem;
    font-weight: 700;
}

.upload-guide strong {
    color: var(--text-primary);
}

.section-divider {
    height: 1px;
    background: linear-gradient(90deg, rgba(0,0,0,0.06), rgba(0,0,0,0.12), rgba(0,0,0,0.06));
    margin: 16px 0 22px 0;
    border: none;
}

.warning-text {
    color: var(--warning);
    font-weight: 600;
    font-size: 0.95rem;
    margin-top: 10px;
    text-align: left;
    padding: 8px 12px;
    background-color: var(--warning-bg);
    border-radius: var(--radius-md);
    border-left: 4px solid var(--warning);
}

/* 강조 색상: XLS vs XLS data 구분 표시 */
.warning-text .xls {
    color: var(--warning);
    background: rgba(217, 119, 6, 0.12);
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 800;
}
.warning-text .xlsdata {
    color: var(--success);
    background: var(--success-bg);
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 800;
}

.privacy-notice {
    margin-top: 10px;
    padding: 14px 16px;
    border-radius: var(--radius-md);
    background: var(--neutral-100);
    border: 1px solid var(--neutral-200);
    color: var(--text-secondary);
}
.privacy-notice p {
    margin: 0 0 8px 0;
    font-weight: 600;
    color: var(--text-primary);
}
.privacy-notice ul {
    margin: 0;
    padding-left: 18px;
    list-style: disc;
    color: var(--text-secondary);
}
.privacy-notice li {
    margin: 3px 0;
    line-height: 1.5;
}
.privacy-notice .privacy-footnote {
    color: var(--text-muted);
    opacity: 1;
    margin-top: 8px;
}

.results-section {
    padding: 28px 36px 36px;
}

/* 하단 크레딧 푸터 */
.app-footer {
    padding: 16px 36px 24px 36px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    border-top: 1px solid var(--neutral-200);
    background: none;
}
.app-footer .footer-right {
    display: flex;
    align-items: center;
    gap: 12px;
}
.app-footer .credits {
    text-align: right;
    font-size: 0.85rem;
    color: var(--text-secondary);
    background: none;
    padding: 0;
    border-radius: 0;
}
.app-footer .credits a:not(.help-btn) {
    color: var(--text-muted);
    text-decoration: none;
    border-bottom: 1px dashed var(--neutral-400);
}
.app-footer .credits a:not(.help-btn):hover {
    color: var(--text-primary);
    border-bottom-color: var(--neutral-500);
}

/* last updated 표시 */
.app-footer .updated {
    font-size: 0.8rem;
    color: var(--text-muted);
    margin-left: 8px;
}

/* 도움말 버튼 */
.help-btn {
    display: inline-block;
    padding: 6px 12px;
    font-size: 0.85rem;
    line-height: 1;
    border-radius: 999px;
    color: var(--primary);
    background: var(--bg-card);
    border: 1px solid var(--primary);
    text-decoration: none;
    transition: all 0.2s ease;
}
.help-btn:hover {
    color: var(--text-inverse);
    background: var(--primary);
    border-color: var(--primary);
    box-shadow: var(--shadow-sm);
}

.tabs {
    display: flex;
    gap: 0;
    padding: 0;
    background: none;
    border: none;
    border-bottom: 2px solid var(--neutral-200);
    border-radius: 0;
    margin-bottom: 30px;
}

.tab-btn {
    flex: 1;
    min-width: 0;
    padding: 14px 18px;
    background: none;
    border: none;
    border-bottom: 3px solid transparent;
    cursor: pointer;
    font-size: 0.95rem;
    font-weight: 500;
    color: var(--text-secondary);
    transition: all 0.25s ease;
    border-radius: 0;
    position: relative;
    margin-bottom: -2px;
}

.tab-btn.active {
    color: var(--primary);
    background: none;
    box-shadow: none;
    border-bottom-color: var(--primary);
    font-weight: 700;
}

.tab-btn:hover:not(.active) {
    background: none;
    color: var(--text-primary);
    border-bottom-color: var(--neutral-300);
}

.tab-content {
    display: none;
}

.tab-content.active {
    display: block;
}

.tab-content h2 {
    color: var(--text-primary);
    margin-bottom: 22px;
    font-size: 1.45rem;
    font-weight: 600;
}


.subject-averages {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 20px;
}

.subject-item {
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-50) 100%);
    border-radius: var(--radius-lg);
    padding: 22px;
    border: 1px solid var(--neutral-200);
    transition: all 0.25s ease;
    box-shadow: none;
}

.subject-item:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.subject-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
}

.subject-header h3 {
    color: var(--text-primary);
    font-size: 1.15rem;
    font-weight: 600;
}

.credits {
    background: var(--neutral-100);
    color: var(--text-secondary);
    padding: 5px 10px;
    border-radius: 15px;
    border: 1px solid var(--neutral-200);
    font-size: 0.8rem;
    font-weight: 500;
}

.average-score {
    text-align: center;
}

.average-score .score {
    display: block;
    font-size: 2.2rem;
    font-weight: 600;
    color: var(--text-primary);
    margin-bottom: 5px;
}

.average-score .label {
    color: var(--text-secondary);
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.achievement-bars {
    margin-top: 20px;
    padding-top: 20px;
    border-top: 1px solid var(--neutral-200);
}

.achievement-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 8px;
}

.achievement-label {
    width: 25px;
    font-weight: 600;
    color: var(--text-primary);
    font-size: 0.9rem;
    text-align: center;
}

.achievement-bar-container {
    flex: 1;
    height: 20px;
    background: var(--neutral-200);
    border-radius: 10px;
    overflow: hidden;
    position: relative;
}

.achievement-bar-fill {
    height: 100%;
    border-radius: 10px;
    transition: width 0.8s ease;
    min-width: 2px;
}

.achievement-bar:nth-child(1) .achievement-bar-fill { background: linear-gradient(135deg, var(--success), var(--success-light)); }
.achievement-bar:nth-child(2) .achievement-bar-fill { background: linear-gradient(135deg, var(--info), var(--info-light)); }
.achievement-bar:nth-child(3) .achievement-bar-fill { background: linear-gradient(135deg, var(--accent), var(--accent-light)); }
.achievement-bar:nth-child(4) .achievement-bar-fill { background: linear-gradient(135deg, var(--warning), var(--primary-light)); }
.achievement-bar:nth-child(5) .achievement-bar-fill { background: linear-gradient(135deg, var(--primary), var(--primary-dark)); }

.achievement-percentage {
    width: 50px;
    text-align: right;
    font-weight: 500;
    color: var(--text-primary);
    font-size: 0.85rem;
}

.achievement-distribution {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 30px;
}

.distribution-item {
    background: var(--bg-card);
    border-radius: var(--radius-md);
    padding: 25px;
    box-shadow: var(--shadow-sm);
    transition: all 0.25s ease;
}

.distribution-item:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.distribution-item h3 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.3rem;
    font-weight: 500;
}

.distribution-bars {
    display: flex;
    flex-direction: column;
    gap: 15px;
}

.grade-bar {
    display: flex;
    align-items: center;
    gap: 15px;
}

.grade-label {
    width: 30px;
    font-weight: 600;
    color: var(--text-primary);
    font-size: 1.1rem;
}

.bar-container {
    flex: 1;
    height: 25px;
    background: var(--neutral-200);
    border-radius: 15px;
    overflow: hidden;
    position: relative;
}

.bar {
    height: 100%;
    background: linear-gradient(135deg, var(--info) 0%, var(--info-light) 100%);
    border-radius: 15px;
    transition: width 0.8s ease;
    min-width: 2px;
}

.percentage {
    width: 60px;
    text-align: right;
    font-weight: 500;
    color: var(--text-primary);
}

.student-analysis {
    width: 100%;
}

.search-box {
    margin-bottom: 25px;
}

.search-box input {
    width: 100%;
    max-width: 400px;
    padding: 14px 18px;
    border: 1px solid var(--neutral-300);
    border-radius: 14px;
    font-size: 1rem;
    outline: none;
    transition: all 0.25s ease;
    background: var(--bg-card);
}

.search-box input:focus {
    border-color: var(--primary);
    box-shadow: 0 0 0 3px var(--primary-bg);
}

.students-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 18px;
    margin-top: 20px;
}

.student-card {
    background: var(--bg-card);
    border-radius: var(--radius-lg);
    box-shadow: none;
    border: 1px solid var(--neutral-200);
    transition: all 0.25s ease;
    overflow: hidden;
}

.student-card:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.student-card-header {
    background: linear-gradient(135deg, #5F4A8B 0%, #483670 100%);
    color: var(--text-inverse);
    padding: 16px 18px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    border-bottom: none;
}

.student-basic-info {
    display: flex;
    align-items: baseline;
    gap: 10px;
    flex-wrap: wrap;
}

.student-basic-info h4 {
    margin: 0;
    font-size: 1.2rem;
    font-weight: 600;
    white-space: nowrap;
}

.student-number {
    font-size: 0.88rem;
    color: var(--text-secondary);
    opacity: 1;
}

.student-summary {
    display: flex;
    flex-direction: row;
    gap: 6px;
    flex-wrap: wrap;
}

.summary-row {
    display: contents;
}

.summary-metric-inline {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    background: var(--neutral-100);
    border: 1px solid var(--neutral-200);
    padding: 4px 8px;
    border-radius: 6px;
    white-space: nowrap;
    flex-shrink: 0;
}

.summary-metric-inline .metric-label {
    font-size: 0.7rem;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    white-space: nowrap;
}

.summary-metric-inline .metric-value {
    font-size: 0.9rem;
    font-weight: 700;
    white-space: nowrap;
    color: var(--text-primary);
}

.summary-metric {
    text-align: center;
    background: rgba(255, 255, 255, 0.15);
    padding: 8px 12px;
    border-radius: 8px;
    min-width: 70px;
}

.summary-metric .metric-label {
    display: block;
    font-size: 0.7rem;
    opacity: 0.8;
    margin-bottom: 2px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.summary-metric .metric-value {
    display: block;
    font-size: 1.1rem;
    font-weight: 700;
}

.student-subjects {
    padding: 15px 20px;
    max-height: none;
    overflow: visible;
}

.subject-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 0;
    border-bottom: 1px solid var(--neutral-200);
}

.subject-row:last-child {
    border-bottom: none;
}

.subject-row.no-grade {
    opacity: 0.7;
}

.subject-name {
    font-weight: 500;
    color: var(--text-primary);
    flex: 1;
    font-size: 0.9rem;
}

.subject-data {
    display: flex;
    gap: 8px;
    align-items: center;
}

.subject-score {
    font-weight: 600;
    color: var(--primary);
    font-size: 0.85rem;
    min-width: 45px;
    text-align: right;
}

.subject-achievement {
    font-size: 0.8rem;
    padding: 2px 6px;
    border-radius: 3px;
    min-width: 20px;
    text-align: center;
}

.subject-grade {
    font-weight: 500;
    color: var(--text-secondary);
    font-size: 0.8rem;
    min-width: 35px;
    text-align: center;
}

.subject-percentile {
    font-weight: 500;
    color: var(--success);
    font-size: 0.8rem;
    min-width: 40px;
    text-align: right;
}

/* ── 학생 카드: 새 뱃지 레이아웃 ── */
.student-card-title-row {
    display: flex;
    align-items: baseline;
    gap: 10px;
    flex-wrap: wrap;
}

.student-card-name {
    margin: 0;
    font-size: 1.15rem;
    font-weight: 700;
    color: #FFFFFF;
    white-space: nowrap;
}

.student-card-class {
    font-size: 0.8rem;
    color: rgba(255, 255, 255, 0.75);
    white-space: nowrap;
}

.student-card-badges {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}

.card-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 5px 12px;
    background: rgba(255, 255, 255, 0.15);
    border: 1px solid rgba(255, 255, 255, 0.3);
    border-radius: 8px;
    white-space: nowrap;
    flex-shrink: 0;
    transition: border-color 0.2s ease;
}

.card-badge:hover {
    border-color: var(--neutral-300);
}

.card-badge-label {
    font-size: 0.68rem;
    font-weight: 600;
    color: rgba(255, 255, 255, 0.7);
    text-transform: uppercase;
    letter-spacing: 0.3px;
}

.card-badge-value {
    font-size: 0.88rem;
    font-weight: 700;
    color: #FFFFFF;
}

.card-badge-value.primary {
    color: #FFFFFF;
}

.card-badge-value.accent {
    color: #E8E0F4;
}

.card-badge-value small {
    font-size: 0.7rem;
    font-weight: 500;
    color: var(--text-muted);
}

.student-card-footer {
    background: var(--bg-card);
    padding: 15px 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid var(--neutral-200);
}

.grade-subjects-count {
    font-size: 0.85rem;
    color: var(--text-secondary);
}

.view-detail-btn {
    background: var(--bg-card);
    color: var(--text-primary);
    border: 1px solid var(--neutral-300);
    padding: 8px 16px;
    border-radius: var(--radius-sm);
    font-size: 0.85rem;
    cursor: pointer;
    transition: all 0.25s ease;
    font-weight: 500;
}

.view-detail-btn:hover {
    transform: translateY(-1px);
    border-color: var(--primary);
    color: var(--primary-dark);
    box-shadow: var(--shadow-sm);
}

.achievement.A {
    background: var(--success);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.B {
    background: var(--info);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.C {
    background: var(--accent);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.D {
    background: var(--warning);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.E, .achievement.미도달 {
    background: var(--primary);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.score {
    font-weight: 600;
    color: var(--text-primary);
}

.grade {
    text-align: center;
    font-weight: 500;
}

.rank {
    text-align: center;
    font-weight: 500;
    color: var(--text-secondary);
}

.avg-grade {
    text-align: center;
    font-weight: 600;
    color: var(--primary);
    font-size: 1.1rem;
}

.grade-analysis-container {
    display: grid;
    grid-template-columns: 1fr 1fr;
    grid-template-rows: auto auto;
    gap: 20px;
    margin-bottom: 30px;
}

.chart-section {
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-50) 100%);
    border-radius: var(--radius-lg);
    padding: 22px;
    text-align: center;
    border: 1px solid var(--neutral-200);
}

.chart-section h3 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.3rem;
    font-weight: 500;
}

.chart-section canvas {
    max-width: 100%;
    height: 350px !important;
}

.stats-section {
    grid-column: 1 / -1;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-50) 100%);
    border-radius: var(--radius-lg);
    padding: 22px;
    border: 1px solid var(--neutral-200);
}

.stat-item {
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    padding: 20px;
    box-shadow: none;
    border: 1px solid var(--neutral-200);
}

.stat-label {
    color: var(--text-secondary);
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 8px;
}

.stat-value {
    color: var(--text-primary);
    font-size: 2rem;
    font-weight: 600;
}

@media (max-width: 768px) {
    .grade-analysis-container {
        grid-template-columns: 1fr;
        gap: 20px;
    }
    
    .stats-section {
        grid-template-columns: repeat(2, 1fr);
        gap: 15px;
    }
    
    .chart-section {
        padding: 15px;
    }
    
    .stat-item {
        padding: 15px;
    }
    
    .stat-value {
        font-size: 1.5rem;
    }
}

.loading {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 60px;
    color: var(--text-secondary);
}

.spinner {
    width: 50px;
    height: 50px;
    border: 4px solid var(--neutral-200);
    border-top: 4px solid var(--primary);
    border-radius: 50%;
    animation: spin 1s linear infinite;
    margin-bottom: 20px;
}

@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}

.loading p {
    font-size: 1.1rem;
    color: var(--text-secondary);
}

.error-message {
    background: var(--warning-bg);
    color: var(--warning);
    padding: 20px;
    margin: 20px 40px;
    border-radius: var(--radius-md);
    border-left: 5px solid var(--warning);
    font-size: 1rem;
}

.file-list {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 20px;
    margin: 20px 0;
    border: 1px solid var(--neutral-200);
}

.file-list h4 {
    color: var(--text-primary);
    margin-bottom: 15px;
    font-size: 1.1rem;
}

.file-list ul {
    list-style: none;
    padding: 0;
}

.file-list li {
    background: var(--bg-card);
    padding: 10px 15px;
    margin: 8px 0;
    border-radius: var(--radius-sm);
    border: 1px solid var(--neutral-200);
    box-shadow: none;
}

.file-selector-section {
    background: var(--bg-card);
    padding: 20px;
    border-radius: var(--radius-lg);
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 15px;
    border: 1px solid var(--neutral-200);
}

.file-selector-section label {
    color: var(--text-primary);
    font-weight: 500;
    white-space: nowrap;
}

.file-select {
    flex: 1;
    padding: 10px 15px;
    border: 2px solid var(--neutral-300);
    border-radius: var(--radius-sm);
    font-size: 1rem;
    background: var(--bg-card);
    outline: none;
    transition: all 0.25s ease;
}

.file-select:focus {
    border-color: var(--primary);
    box-shadow: 0 0 0 3px var(--primary-bg);
}

.comparison-container {
    display: grid;
    grid-template-columns: 1fr;
    gap: 30px;
}

.comparison-section {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 25px;
    border: 1px solid var(--border-light);
}

.comparison-section h3 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.3rem;
    font-weight: 500;
}

.comparison-table {
    width: 100%;
    border-collapse: collapse;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    overflow: hidden;
    box-shadow: var(--shadow-sm);
}

.comparison-table th,
.comparison-table td {
    padding: 12px 15px;
    text-align: center;
    border-bottom: 1px solid var(--neutral-200);
}

.comparison-table th {
    background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
    color: var(--text-inverse);
    font-weight: 500;
    font-size: 0.9rem;
}

.comparison-table tr:nth-child(even) {
    background: var(--neutral-100);
}

.comparison-table tr:hover {
    background: var(--primary-bg);
}

@media (max-width: 768px) {
    .file-selector-section {
        flex-direction: column;
        align-items: stretch;
        gap: 10px;
    }
    
    .file-select {
        width: 100%;
    }
    
    .comparison-table {
        font-size: 0.8rem;
    }
    
    .comparison-table th,
    .comparison-table td {
        padding: 8px 6px;
    }
}

/* 학생 선택 및 상세 분석 스타일 */
.student-selector {
    display: flex;
    align-items: center;
    gap: 15px;
    background: var(--bg-card);
    padding: 20px;
    border-radius: var(--radius-lg);
    margin-bottom: 20px;
    flex-wrap: wrap;
    border: 1px solid var(--neutral-200);
    box-shadow: var(--shadow-sm);
}

.selector-group {
    display: flex;
    align-items: center;
    gap: 8px;
}

.selector-group label {
    font-weight: 500;
    color: var(--text-primary);
    white-space: nowrap;
}

.selector {
    padding: 10px 12px;
    border: 1px solid var(--neutral-300);
    border-radius: var(--radius-sm);
    font-size: 0.9rem;
    background: var(--bg-card);
    min-width: 120px;
    transition: all 0.25s ease;
}

.selector:focus {
    border-color: var(--primary);
    outline: none;
    box-shadow: 0 0 0 3px var(--primary-bg);
}

.detail-btn {
    background: var(--bg-card);
    color: var(--text-primary);
    border: 1px solid var(--neutral-300);
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    font-size: 0.9rem;
    cursor: pointer;
    transition: all 0.25s ease;
    white-space: nowrap;
}

.detail-btn:hover:not(:disabled) {
    transform: translateY(-2px);
    box-shadow: var(--shadow-sm);
    border-color: var(--primary);
    color: var(--primary-dark);
}

.detail-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}

.view-toggle {
    display: flex;
    background: var(--neutral-100);
    border-radius: var(--radius-md);
    padding: 4px;
    margin-bottom: 20px;
    width: fit-content;
    border: 1px solid var(--neutral-200);
}

.toggle-btn {
    background: none;
    border: none;
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    cursor: pointer;
    transition: all 0.25s ease;
    font-weight: 500;
    color: var(--text-secondary);
}

.toggle-btn.active {
    background: var(--bg-card);
    color: var(--text-primary);
    box-shadow: var(--shadow-sm);
}

/* 학생 상세 분석 스타일 */
.student-detail-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-100) 100%);
    color: var(--text-primary);
    padding: 22px 24px;
    border-radius: var(--radius-lg);
    margin-bottom: 20px;
    border: 1px solid var(--border-light);
    box-shadow: var(--shadow-md);
}

.student-info h3 {
    font-size: 1.5rem;
    margin-bottom: 6px;
    font-weight: 400;
}

.student-meta {
    display: flex;
    gap: 14px;
    font-size: 0.85rem;
    opacity: 0.9;
    flex-wrap: wrap;
}

.overall-stats {
    display: flex;
    gap: 12px;
}

.stat-card {
    text-align: center;
    background: var(--bg-card);
    padding: 12px 16px;
    border-radius: var(--radius-md);
    border: 1px solid var(--border-light);
    box-shadow: var(--shadow-sm);
    min-width: 110px;
}

.stat-label {
    display: block;
    font-size: 0.8rem;
    opacity: 0.8;
    margin-bottom: 5px;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.stat-value {
    display: block;
    font-size: 1.35rem;
    font-weight: 700;
    color: var(--text-primary);
}

.stat-value.grade {
    color: var(--primary);
}

.student-detail-content {
    display: flex;
    flex-direction: column;
    gap: 22px;
    margin-bottom: 24px;
}

.analysis-overview {
    display: grid;
    grid-template-columns: minmax(0, 1.35fr) minmax(280px, 0.85fr);
    gap: 20px;
    margin-bottom: 20px;
    align-items: start;
}

.student-summary {
    display: flex;
    flex-direction: column;
    gap: 20px;
}

.summary-card {
    background: var(--bg-card);
    border-radius: var(--radius-lg);
    padding: 18px 20px;
    box-shadow: var(--shadow-md);
    border: 1px solid var(--border-light);
}

.summary-header {
    margin-bottom: 14px;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--neutral-200);
}

.summary-header h4 {
    color: var(--text-primary);
    font-size: 1.2rem;
    font-weight: 600;
    margin: 0;
}

.summary-grid {
    display: grid;
    gap: 10px;
}

.summary-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 0;
    border-bottom: 1px solid var(--border-light);
}

.summary-item:last-child {
    border-bottom: none;
}

.summary-label {
    font-weight: 500;
    color: var(--text-secondary);
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.summary-value {
    font-weight: 600;
    color: var(--text-primary);
    font-size: 1rem;
    text-align: right;
}

.summary-value-group {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 4px;
}

.summary-value.highlight {
    color: var(--primary);
    font-size: 1.05rem;
    font-weight: 600;
}

.summary-value.orange {
    color: var(--accent);
    font-size: 1.05rem;
    font-weight: 600;
}

.summary-note {
    font-size: 0.72rem;
    color: var(--text-muted);
    text-align: right;
    line-height: 1.35;
    white-space: nowrap;
}

.metric-value.orange {
    color: var(--accent);
    font-weight: 600;
}

.chart-container {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 16px 18px 18px;
    text-align: center;
    border: 1px solid var(--border-light);
    width: 100%;
    max-width: 420px;
    justify-self: end;
}

.chart-container h4 {
    color: var(--text-primary);
    margin-bottom: 12px;
    font-size: 1rem;
    font-weight: 600;
}

.chart-container canvas {
    display: block;
    width: min(100%, 320px) !important;
    height: auto !important;
    margin: 0 auto;
}

.subject-details h4 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.2rem;
    font-weight: 500;
}

.subject-cards {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 16px;
    max-height: none;
}

@media (max-width: 1200px) {
    .subject-cards {
        grid-template-columns: 1fr;
    }
}

/* 교과(군)별 섹션 스타일 */
.subject-group-section {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 20px;
    border: 1px solid var(--border-light);
}

.subject-group-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 12px 16px;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    margin-bottom: 16px;
    box-shadow: var(--shadow-sm);
}

.subject-group-header h5 {
    margin: 0;
    font-size: 1rem;
    font-weight: 600;
    color: var(--text-primary);
}

.subject-group-header .subject-count {
    font-size: 0.8rem;
    color: var(--text-secondary);
    background: var(--neutral-200);
    padding: 4px 10px;
    border-radius: 10px;
    font-weight: 500;
}

.subject-group-cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
    gap: 15px;
}

/* 컴팩트 테이블 스타일 */
.subject-group-section.compact {
    padding: 14px;
    margin-bottom: 0;
    height: fit-content;
}

.subject-group-section.compact .subject-group-header {
    margin-bottom: 12px;
    padding: 8px 12px;
}

.subject-group-section.compact .subject-group-header h5 {
    font-size: 0.95rem;
}

.subject-table {
    width: 100%;
    border-collapse: collapse;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    overflow: hidden;
    font-size: 0.8rem;
}

.subject-table thead {
    background: linear-gradient(135deg, var(--neutral-200) 0%, var(--neutral-100) 100%);
}

.subject-table th {
    padding: 8px 6px;
    text-align: center;
    font-weight: 600;
    color: var(--text-primary);
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.3px;
    border-bottom: 2px solid var(--neutral-300);
}

.subject-table th:first-child {
    text-align: left;
    padding-left: 10px;
}

.subject-table td {
    padding: 8px 6px;
    border-bottom: 1px solid var(--neutral-200);
    color: var(--text-primary);
}

.subject-table td.center {
    text-align: center;
}

.subject-table td.subject-name-cell {
    font-weight: 500;
    padding-left: 10px;
    max-width: 100px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

.subject-table tbody tr:hover {
    background: var(--neutral-100);
}

.subject-table tbody tr:last-child td {
    border-bottom: none;
}

.subject-table tr.no-grade-row {
    opacity: 0.7;
    background: var(--neutral-50);
}

.subject-table .score-value {
    font-weight: 600;
    color: var(--text-primary);
}

.subject-table .avg-value {
    font-size: 0.75rem;
    color: var(--text-muted);
    margin-left: 2px;
}

.subject-table .achievement-badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 600;
    font-size: 0.8rem;
}

.subject-table .achievement-badge.A { background: var(--success); color: var(--text-inverse); }
.subject-table .achievement-badge.B { background: var(--info); color: var(--text-inverse); }
.subject-table .achievement-badge.C { background: var(--accent); color: var(--text-inverse); }
.subject-table .achievement-badge.D { background: var(--warning); color: var(--text-inverse); }
.subject-table .achievement-badge.E { background: var(--primary); color: var(--text-inverse); }

.subject-table .grade9-value {
    color: var(--accent);
    font-weight: 600;
}

@media (max-width: 768px) {
    .subject-table {
        font-size: 0.75rem;
    }

    .subject-table th,
    .subject-table td {
        padding: 8px 4px;
    }

    .subject-table td.subject-name-cell {
        max-width: 80px;
    }

    .subject-table .avg-value {
        display: none;
    }
}

.subject-card.no-grade {
    opacity: 0.8;
    border-left: 4px solid var(--neutral-500);
}

.subject-metrics.simple {
    grid-template-columns: 1fr 1fr;
    margin-bottom: 0;
}

.no-grade-notice {
    text-align: center;
    padding: 15px;
    background: var(--neutral-200);
    border-radius: var(--radius-sm);
    margin-top: 15px;
}

.no-grade-notice span {
    color: var(--text-secondary);
    font-size: 0.9rem;
    font-style: italic;
}

.subject-card {
    background: var(--bg-card);
    border-radius: var(--radius-md);
    padding: 18px;
    box-shadow: var(--shadow-sm);
    transition: all 0.2s ease;
    border: 1px solid var(--border-light);
}

.subject-card:hover {
    box-shadow: var(--shadow-md);
    border-color: var(--border-medium);
}

.subject-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
    padding-bottom: 10px;
    border-bottom: 2px solid var(--neutral-200);
}

.subject-header h5 {
    color: var(--text-primary);
    font-size: 1.1rem;
    font-weight: 600;
    margin: 0;
}

.subject-header .credits {
    background: var(--info);
    color: var(--text-inverse);
    padding: 4px 8px;
    border-radius: 12px;
    font-size: 0.8rem;
    font-weight: 500;
}

.subject-metrics {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 15px;
    margin-bottom: 15px;
}

.subject-metrics:last-of-type {
    grid-template-columns: 1fr 1fr 0fr;
}

.metric {
    text-align: center;
}

.metric-label {
    display: block;
    font-size: 0.8rem;
    color: var(--text-secondary);
    margin-bottom: 5px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.metric-value {
    display: block;
    font-size: 1rem;
    font-weight: 600;
    color: var(--text-primary);
}

.metric-average {
    display: block;
    font-size: 0.8rem;
    color: var(--text-secondary);
    font-weight: normal;
    margin-top: 2px;
}

.metric-value.achievement.A {
    background: var(--success);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.B {
    background: var(--info);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.C {
    background: var(--accent);
    color: var(--text-primary);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.D {
    background: var(--warning);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.E, .metric-value.achievement.미도달 {
    background: var(--primary);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.percentile-bar {
    height: 8px;
    background: var(--neutral-200);
    border-radius: 4px;
    overflow: hidden;
    position: relative;
}

.percentile-fill {
    height: 100%;
    border-radius: 4px;
    transition: width 0.8s ease;
}

.percentile-fill.excellent { background: linear-gradient(90deg, var(--success), var(--success-light)); }
.percentile-fill.good { background: linear-gradient(90deg, var(--info), var(--info-light)); }
.percentile-fill.average { background: linear-gradient(90deg, var(--warning), var(--warning-light)); }
.percentile-fill.low { background: linear-gradient(90deg, var(--neutral-500), var(--neutral-400)); }

.percentile.excellent { color: var(--success); font-weight: 600; }
.percentile.good { color: var(--info); font-weight: 600; }
.percentile.average { color: var(--warning); font-weight: 600; }
.percentile.low { color: var(--neutral-500); font-weight: 500; }

@media (max-width: 1024px) {
    .analysis-overview {
        grid-template-columns: 1fr;
        gap: 20px;
    }
    
    .chart-container {
        padding: 20px;
    }
    
    .student-detail-header {
        flex-direction: column;
        gap: 20px;
    }
    
    .overall-stats {
        align-self: stretch;
        justify-content: space-around;
    }
    
    .subject-cards {
        grid-template-columns: 1fr;
    }
}

/* 출력용 스타일 */
@page {
    size: A4 portrait;
    margin: 10mm;
}
@media print {
    .print-area {
        transform-origin: top left !important;
    }
    .print-area.apply-print-scale {
        transform: scale(var(--page-scale, 1)) !important;
    }
    * {
        -webkit-print-color-adjust: exact !important;
        color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    body {
        background: white !important;
        margin: 0;
        padding: 0;
        font-size: 12px;
        line-height: 1.4;
        color: #000 !important;
    }
    
    .container {
        max-width: none;
        margin: 0;
        box-shadow: none;
        border-radius: 0;
        background: white;
    }
    
    header {
        display: none !important;
    }
    
    .upload-section,
    .tabs,
    .view-toggle,
    .student-selector,
    .search-box,
    .print-controls {
        display: none !important;
    }
    
    .results-section {
        padding: 15px;
    }
    
    .tab-content {
        display: block !important;
    }
    
    .tab-content:not(.print-target) {
        display: none !important;
    }

    /* 학생 탭 인쇄 시 개인 상세 페이지만 표시 */
    #students-tab.only-class-print > *:not(.class-print-area) {
        display: none !important;
    }

    /* A4에 맞춘 폭 고정 및 중앙 정렬 */
    .class-print-area {
        width: 190mm;
        margin: 0 auto;
    }
    .class-print-area .student-print-page {
        width: 190mm;
        transform-origin: top left !important;
    }
    .class-print-area .student-print-page.apply-print-scale {
        transform: scale(var(--page-scale, 1)) !important;
    }

    /* 학급 전체 인쇄 모드: 더 컴팩트한 카드와 차트 크기 */
    #students-tab.only-class-print .student-detail-header {
        padding: 12px;
        margin-bottom: 10px;
    }
    #students-tab.only-class-print .student-info h3 {
        font-size: 14px;
    }
    #students-tab.only-class-print .student-meta {
        font-size: 11px;
    }
    #students-tab.only-class-print .summary-card,
    #students-tab.only-class-print .stat-card {
        margin-bottom: 8px;
        padding: 10px;
    }
    
    /* 별도 프린트 헤더는 사용하지 않음 */
    .print-header {
        display: none !important;
    }
    
    .print-header h2 {
        margin: 0;
        color: #2c3e50;
        font-size: 18px;
        font-weight: bold;
    }
    
    .print-date {
        margin-top: 10px;
        font-size: 12px;
        color: #666;
    }
    
    .student-detail-header {
        background: #f8f9fa !important;
        border: 2px solid #4facfe;
        margin-bottom: 15px;
        padding: 20px;
        page-break-after: avoid;
    }
    
    .student-info h3 {
        color: #2c3e50 !important;
        font-size: 16px;
        margin-bottom: 8px;
    }
    
    .student-meta {
        color: #666 !important;
        font-size: 12px;
    }
    
    .stat-card {
        border: 1px solid #ddd !important;
        background: white !important;
    }
    
    .stat-label {
        color: #666 !important;
        font-size: 10px;
    }
    
    .stat-value {
        color: #2c3e50 !important;
        font-size: 14px;
    }
    
    .analysis-overview {
        grid-template-columns: 1fr;
        gap: 15px;
        page-break-inside: avoid;
    }
    
    /* 레이더 차트도 출력/PDF에 포함 */
    .chart-container {
        display: block !important;
    }
    
    .summary-card {
        border: 1px solid #ddd !important;
        background: #f9f9f9 !important;
        margin-bottom: 15px;
    }
    
    .summary-header h4 {
        color: #2c3e50 !important;
        font-size: 14px;
    }
    
    .summary-label {
        color: #666 !important;
        font-size: 11px;
    }
    
    .summary-value {
        color: #2c3e50 !important;
        font-size: 12px;
    }
    
    .summary-value.highlight {
        color: #4facfe !important;
        font-weight: bold;
    }
    
    .subject-details h4 {
        color: #2c3e50 !important;
        font-size: 14px;
        margin-bottom: 15px;
    }
    
    .subject-cards {
        grid-template-columns: repeat(2, 1fr);
        gap: 10px;
        page-break-inside: avoid;
    }
    
    .subject-card {
        border: 1px solid #ddd !important;
        background: white !important;
        page-break-inside: avoid;
        margin-bottom: 8px;
        padding: 12px;
    }
    
    .subject-header h5 {
        color: #2c3e50 !important;
        font-size: 12px;
        margin: 0 0 8px 0;
    }
    
    .subject-header .credits {
        background: #4facfe !important;
        color: white !important;
        font-size: 9px;
        padding: 2px 6px;
    }
    
    .subject-metrics {
        gap: 8px;
        margin-bottom: 8px;
    }
    
    .metric-label {
        font-size: 9px;
        color: #666 !important;
    }
    
    .metric-value {
        font-size: 11px;
        color: #2c3e50 !important;
    }
    
    .metric-average {
        font-size: 9px;
        color: #666 !important;
    }
    
    .percentile-bar {
        height: 6px;
        background: #e9ecef !important;
    }
    
    .no-grade-notice {
        background: rgba(108, 117, 125, 0.1) !important;
        font-size: 10px;
    }
    
    .no-grade-notice span {
        color: #666 !important;
    }
    
    /* 성취도 색상 */
    .achievement.A, .metric-value.achievement.A { 
        background: #28a745 !important; 
        color: white !important; 
    }
    .achievement.B, .metric-value.achievement.B { 
        background: #17a2b8 !important; 
        color: white !important; 
    }
    .achievement.C, .metric-value.achievement.C { 
        background: #ffc107 !important; 
        color: #212529 !important; 
    }
    .achievement.D, .metric-value.achievement.D { 
        background: #fd7e14 !important; 
        color: white !important; 
    }
    .achievement.E, .achievement.미도달, 
    .metric-value.achievement.E, .metric-value.achievement.미도달 { 
        background: #dc3545 !important; 
        color: white !important; 
    }
    
    /* 페이지 나누기 규칙 */
    .subject-card {
        break-inside: avoid;
    }
    
    .summary-card {
        break-inside: avoid;
    }

    /* 학급 전체 인쇄: 학생별 한 페이지씩 */
    .class-print-area .student-print-page {
        page-break-after: always;
        break-after: page;
    }
    .class-print-area .student-print-page:last-child {
        page-break-after: auto;
        break-after: auto;
    }
}

/* PDF 출력 버튼 스타일 */
.print-controls {
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
}

.student-nav-controls {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
}

.student-nav-status {
    color: var(--text-secondary);
    font-size: 0.9rem;
    font-weight: 600;
    padding: 0 4px;
}

.print-btn, .pdf-btn {
    background: linear-gradient(135deg, var(--success) 0%, var(--success-light) 100%);
    color: var(--text-inverse);
    border: none;
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    font-size: 0.9rem;
    cursor: pointer;
    transition: all 0.25s ease;
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 500;
}

.pdf-btn {
    background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
}

.print-btn:hover, .pdf-btn:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.print-btn::before {
    content: "🖨️";
    font-size: 16px;
}

.pdf-btn::before {
    content: "📄";
    font-size: 16px;
}

@media (max-width: 768px) {
    .student-selector {
        flex-direction: column;
        align-items: stretch;
        gap: 15px;
    }

    .selector-group {
        justify-content: space-between;
    }

    .selector {
        min-width: unset;
        flex: 1;
    }

    .subject-metrics {
        grid-template-columns: repeat(2, 1fr);
        gap: 10px;
    }

    .container {
        margin: 10px;
        border-radius: var(--radius-lg);
    }

    header {
        padding: 20px;
    }

    header h1 {
        font-size: 1.6rem;
    }

    .header-subtitle {
        font-size: 0.92rem;
    }

    .upload-section,
    .results-section {
        padding: 20px;
    }

    .file-input-label {
        min-height: 74px;
        padding: 16px 18px;
    }

    .action-buttons {
        flex-direction: column;
        align-items: stretch;
    }

    .subject-averages {
        grid-template-columns: 1fr;
    }

    .tabs {
        display: flex;
        width: 100%;
    }

    .tab-btn {
        min-width: 0;
        padding: 12px 10px;
        font-size: 0.85rem;
    }

    .students-grid {
        grid-template-columns: 1fr;
        gap: 15px;
    }

    .student-card-header {
        padding: 14px 16px;
        gap: 8px;
    }

    .student-card-badges {
        gap: 6px;
    }

    .card-badge {
        padding: 4px 8px;
        gap: 4px;
    }

    .card-badge-label {
        font-size: 0.6rem;
    }

    .card-badge-value {
        font-size: 0.8rem;
    }

    .student-summary {
        gap: 5px;
    }

    .summary-metric-inline {
        padding: 3px 6px;
    }

    .summary-metric-inline .metric-label {
        font-size: 0.6rem;
    }

    .summary-metric-inline .metric-value {
        font-size: 0.78rem;
    }

    .subject-data {
        gap: 6px;
    }

    .subject-name {
        font-size: 0.85rem;
    }

    .subject-score, .subject-achievement, .subject-grade, .subject-percentile {
        font-size: 0.75rem;
    }

    .print-controls {
        justify-content: center;
    }

    .print-btn, .pdf-btn {
        flex: 1;
        min-width: 120px;
        justify-content: center;
    }

    .analyze-btn,
    .secondary-btn {
        width: 100%;
        justify-content: center;
    }

    .print-controls,
    .student-nav-controls {
        justify-content: center;
    }
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
            const introHeader = document.querySelector('.container > header');
            if (introHeader) introHeader.style.display = 'none';
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
        if (document.querySelector('[data-tab="subjects"]') && document.getElementById('subjects-tab')) {
            this.switchTab('subjects');
        }
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

    getStudentDetailNavigationStudents() {
        if (!this.combinedData) return [];

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');
        const studentSearch = document.getElementById('studentSearch');

        const selectedGrade = gradeSelect ? gradeSelect.value : '';
        const selectedClass = classSelect ? classSelect.value : '';
        const detailQuery = studentNameSearch && studentNameSearch.value
            ? studentNameSearch.value.trim().toLowerCase()
            : '';
        const tableQuery = studentSearch && studentSearch.value
            ? studentSearch.value.trim().toLowerCase()
            : '';

        let students = this.combinedData.students;
        if (selectedGrade) {
            students = students.filter(s => String(s.grade) === String(selectedGrade));
        }
        if (selectedClass) {
            students = students.filter(s => String(s.class) === String(selectedClass));
        }
        if (detailQuery) {
            students = students.filter(s =>
                (s.name && s.name.toLowerCase().includes(detailQuery)) ||
                (s.originalNumber && String(s.originalNumber).includes(detailQuery))
            );
        }
        if (tableQuery) {
            students = students.filter(s =>
                (s.name && s.name.toLowerCase().includes(tableQuery)) ||
                String(s.number).includes(tableQuery) ||
                (s.originalNumber && String(s.originalNumber).includes(tableQuery))
            );
        }

        return students;
    }

    navigateStudentDetail(offset) {
        const studentSelect = document.getElementById('studentSelect');
        const currentStudentId = studentSelect ? studentSelect.value : '';
        const navigationStudents = this.getStudentDetailNavigationStudents();
        if (!currentStudentId || navigationStudents.length === 0) return;

        const currentIndex = navigationStudents.findIndex(student => String(student.number) === String(currentStudentId));
        if (currentIndex === -1) return;

        const nextIndex = currentIndex + offset;
        if (nextIndex < 0 || nextIndex >= navigationStudents.length) return;

        const targetStudent = navigationStudents[nextIndex];
        if (studentSelect) {
            studentSelect.value = targetStudent.number;
        }
        this.renderStudentDetail(targetStudent);
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
            const hasRankBasedPercentile = student.ranks 
                && Object.values(student.ranks).some(r => r !== undefined && r !== null && !isNaN(r));
            
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
                    <div class="student-card-title-row">
                        <h4 class="student-card-name">${student.name}</h4>
                        <span class="student-card-class">${student.grade}학년 ${student.class}반 ${student.originalNumber}번</span>
                    </div>
                    <div class="student-card-badges">
                        <div class="card-badge">
                            <span class="card-badge-label">평균등급</span>
                            <span class="card-badge-value primary">${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                        </div>
                        ${averageGradeRank !== null && averageGradeRank !== undefined ? `
                        <div class="card-badge">
                            <span class="card-badge-label">등급순위</span>
                            <span class="card-badge-value">${averageGradeRank}/${totalGradedStudents}위${sameGradeCount > 1 ? ` (${sameGradeCount}명)` : ''}</span>
                        </div>
                        ` : ''}
                        ${weightedAveragePercentile ? `
                        <div class="card-badge">
                            <span class="card-badge-label">백분위</span>
                            <span class="card-badge-value accent">${weightedAveragePercentile.toFixed(1)}%${!hasRankBasedPercentile ? ' <small>(추정)</small>' : ''}</span>
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

    filterStudentTable() {
        if (!this.combinedData) return;

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSearch = document.getElementById('studentSearch');

        const selectedGrade = gradeSelect ? gradeSelect.value : '';
        const selectedClass = classSelect ? classSelect.value : '';
        const searchTerm = studentSearch ? studentSearch.value.trim().toLowerCase() : '';

        // 학년/반/검색어로 필터링
        let filtered = this.combinedData.students;

        if (selectedGrade) {
            filtered = filtered.filter(s => String(s.grade) === String(selectedGrade));
        }

        if (selectedClass) {
            filtered = filtered.filter(s => String(s.class) === String(selectedClass));
        }

        if (searchTerm) {
            filtered = filtered.filter(s =>
                s.number.toString().includes(searchTerm) ||
                s.name.toLowerCase().includes(searchTerm)
            );
        }

        // 테이블 다시 렌더링
        const container = document.getElementById('studentTable');
        if (container) {
            this.renderStudentTable(filtered, this.combinedData.subjects, container);
        }
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
    
        // ★ 추가: 석차 기반 백분위인지, 등급 기반 추정인지 판별
        const hasRankBasedPercentile = student.ranks 
            && Object.values(student.ranks).some(r => r !== undefined && r !== null && !isNaN(r));
        const percentileEstimateNote = (weightedAveragePercentile && !hasRankBasedPercentile)
            ? '<span class="summary-note">등급 기반 추정치</span>'
            : '';
    
        // 평균등급 기준 순위
        const averageGradeRank = student.averageGradeRank;
        const sameGradeCount = student.sameGradeCount;
        const totalGradedStudents = student.totalGradedStudents;
        const studentSelect = document.getElementById('studentSelect');
        if (studentSelect) {
            studentSelect.value = student.number;
        }

        const filteredStudents = this.getStudentDetailNavigationStudents();
        const navigationStudents = filteredStudents.some(s => String(s.number) === String(student.number))
            ? filteredStudents
            : [student];
        const navigationIndex = navigationStudents.findIndex(s => String(s.number) === String(student.number));
        const hasPrevStudent = navigationIndex > 0;
        const hasNextStudent = navigationIndex >= 0 && navigationIndex < navigationStudents.length - 1;
        const navigationLabel = navigationStudents.length > 0 && navigationIndex >= 0
            ? `${navigationIndex + 1} / ${navigationStudents.length}`
            : '';
        const usesBusanReference = this.usesBusanNineGradeReference(student, this.combinedData.subjects);
        const nineGradeReferenceNote = usesBusanReference
            ? '<span class="summary-note">부산교육청 환산 기준</span>'
            : '';
        
        const html = `
            <div class="print-controls">
                <div class="student-nav-controls">
                    <button class="detail-btn student-nav-btn" data-nav-offset="-1" ${hasPrevStudent ? '' : 'disabled'}>이전 학생</button>
                    <span class="student-nav-status">${navigationLabel}</span>
                    <button class="detail-btn student-nav-btn" data-nav-offset="1" ${hasNextStudent ? '' : 'disabled'}>다음 학생</button>
                </div>
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
                            ${student.fileName ? `<span class="file-info">출처: ${student.fileName}</span>` : ''}
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
                                        <span class="summary-value-group">
                                            <span class="summary-value orange">${student.weightedAverage9Grade ? student.weightedAverage9Grade.toFixed(2) : 'N/A'}</span>
                                            ${nineGradeReferenceNote}
                                        </span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">등급 순위</span>
                                        <span class="summary-value highlight">${averageGradeRank !== null && averageGradeRank !== undefined ? `${averageGradeRank}/${totalGradedStudents}위` + (sameGradeCount > 1 ? ` (${sameGradeCount}명)` : '') : 'N/A'}</span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">과목평균 백분위</span>
                                        <span class="summary-value-group">
                                            <span class="summary-value highlight">${weightedAveragePercentile ? weightedAveragePercentile.toFixed(1) + '%' : 'N/A'}</span>
                                            ${percentileEstimateNote}
                                        </span>
                                    </div>
                                    <div class="summary-item">
                                        <span class="summary-label">전체 학생수</span>
                                        <span class="summary-value">${student.totalStudents || 'N/A'}명</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="chart-container">
                            <h4>교과(군)별 평균등급</h4>
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

        container.querySelectorAll('.student-nav-btn').forEach(button => {
            button.addEventListener('click', () => {
                const offset = parseInt(button.getAttribute('data-nav-offset'), 10);
                if (!isNaN(offset)) {
                    this.navigateStudentDetail(offset);
                }
            });
        });
        
        // 레이더 차트 생성
        setTimeout(() => {
            this.createStudentPercentileChart(student);
        }, 100);
    }

    // 학급 전체 인쇄용: 개별 학생과 완전히 동일한 HTML 구조
    buildStudentDetailHTMLForPrint(student, canvasId) {
        const weightedAveragePercentile = this.calculateWeightedAveragePercentile(student, this.combinedData.subjects);
    
        // ★ 추가
        const hasRankBasedPercentile = student.ranks 
            && Object.values(student.ranks).some(r => r !== undefined && r !== null && !isNaN(r));
        const percentileEstimateNote = (weightedAveragePercentile && !hasRankBasedPercentile)
            ? '<span class="summary-note">등급 기반 추정치</span>'
            : '';
    
        const averageGradeRank = student.averageGradeRank;
        const sameGradeCount = student.sameGradeCount;
        const totalGradedStudents = student.totalGradedStudents;
        const usesBusanReference = this.usesBusanNineGradeReference(student, this.combinedData.subjects);
        const nineGradeReferenceNote = usesBusanReference
            ? '<span class="summary-note">부산교육청 환산 기준</span>'
            : '';
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
                                ${student.fileName ? `<span class="file-info">출처: ${student.fileName}</span>` : ''}
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
                                            <span class="summary-value-group">
                                                <span class="summary-value orange">${student.weightedAverage9Grade ? student.weightedAverage9Grade.toFixed(2) : 'N/A'}</span>
                                                ${nineGradeReferenceNote}
                                            </span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">등급 순위</span>
                                            <span class="summary-value highlight">${averageGradeRank !== null && averageGradeRank !== undefined ? `${averageGradeRank}/${totalGradedStudents}위` + (sameGradeCount > 1 ? ` (${sameGradeCount}명)` : '') : 'N/A'}</span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">과목평균 백분위</span>
                                            <span class="summary-value-group">
                                                <span class="summary-value highlight">${weightedAveragePercentile ? weightedAveragePercentile.toFixed(1) + '%' : 'N/A'}</span>
                                                ${percentileEstimateNote}
                                            </span>
                                        </div>
                                        <div class="summary-item">
                                            <span class="summary-label">전체 학생수</span>
                                            <span class="summary-value">${student.totalStudents || 'N/A'}명</span>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="chart-container">
                                <h4>교과(군)별 평균등급</h4>
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

    // 다중 생성용 차트 (교과군별 - PDF용)
    createStudentPercentileChartFor(canvas, student) {
        if (!canvas) return null;

        // 교과군별 평균 등급 계산
        const groupGrades = this.calculateGroupGrades(student);
        if (Object.keys(groupGrades).length === 0) return null;

        // order 순으로 정렬
        const sortedGroups = Object.entries(groupGrades)
            .sort((a, b) => a[1].order - b[1].order);

        const labels = sortedGroups.map(([name]) => name);
        const gradeData = sortedGroups.map(([name, data]) => 6 - data.averageGrade);
        const colors = sortedGroups.map(([name, data]) => data.color);
        const originalGrades = sortedGroups.map(([name, data]) => data.averageGrade);

        // 기존 차트 인스턴스가 해당 캔버스에 남아있다면 파괴
        try {
            const existing = (Chart.getChart ? Chart.getChart(canvas) : (canvas && (canvas._chart || canvas.chart)));
            if (existing && typeof existing.destroy === 'function') existing.destroy();
        } catch (_) {}

        return new Chart(canvas, {
            type: 'radar',
            plugins: [ChartDataLabels],
            data: {
                labels,
                datasets: [{
                    label: '교과군별 평균등급',
                    data: gradeData,
                    backgroundColor: 'rgba(52, 152, 219, 0.2)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 2,
                    pointBackgroundColor: colors,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointRadius: 8
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
                            const idx = context.dataIndex;
                            return originalGrades[idx].toFixed(2) + '등급';
                        },
                        color: '#2c3e50',
                        backgroundColor: 'rgba(255, 255, 255, 0.9)',
                        borderColor: function(context) {
                            return colors[context.dataIndex];
                        },
                        borderWidth: 2,
                        borderRadius: 6,
                        padding: 6,
                        font: {
                            size: 11,
                            weight: 'bold'
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
                                size: 13,
                                weight: '600'
                            },
                            color: function(context) {
                                return colors[context.index] || '#2c3e50';
                            }
                        }
                    }
                }
            }
        });
    }


    // 학급 전체 PDF
    async generateSelectedClassPDF() {
        if (this._pdfGenerating) return; // 중복 클릭 방지
        this._pdfGenerating = true;
        const pdfBtn = document.getElementById('pdfClassBtn');
        const prevBtnHTML = pdfBtn ? pdfBtn.innerHTML : '';
        if (pdfBtn) {
            pdfBtn.disabled = true;
            pdfBtn.innerText = '학급 PDF 생성 중...';
        }
        this.showPdfOverlay();
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

            const total = students.length;
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

                // 진행률 업데이트
                this.updatePdfProgress(i + 1, total);
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
                    let processed = 0;
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

                            // 진행률 업데이트 (분할 저장에서도 누적 기준)
                            processed += 1;
                            this.updatePdfProgress(processed, students.length);
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
        } finally {
            // UI 복구
            this.hidePdfOverlay();
            if (pdfBtn) {
                pdfBtn.disabled = false;
                pdfBtn.innerHTML = prevBtnHTML || '학급 전체 PDF';
            }
            this._pdfGenerating = false;
        }
    }

    showPdfOverlay() {
        try {
            let overlay = document.getElementById('pdfOverlay');
            if (!overlay) {
                overlay = document.createElement('div');
                overlay.id = 'pdfOverlay';
                overlay.style.position = 'fixed';
                overlay.style.left = '0';
                overlay.style.top = '0';
                overlay.style.right = '0';
                overlay.style.bottom = '0';
                overlay.style.background = 'rgba(255,255,255,0.65)';
                overlay.style.zIndex = '9999';
                overlay.style.display = 'flex';
                overlay.style.alignItems = 'center';
                overlay.style.justifyContent = 'center';
                overlay.innerHTML = '<div style="text-align:center;min-width:260px">\
<div class="spinner" style="margin:0 auto 12px auto"></div>\
<div id="pdfOverlayText" style="margin-bottom:10px">학급 PDF 생성 중...</div>\
<div style="height:10px;background:#e9ecef;border-radius:6px;overflow:hidden">\
  <div id="pdfOverlayBar" style="height:100%;width:0%;background:#4facfe;transition:width .2s ease"></div>\
</div>\
</div>';
                document.body.appendChild(overlay);
            } else {
                overlay.style.display = 'flex';
            }
        } catch (_) {}
    }

    hidePdfOverlay() {
        try {
            const overlay = document.getElementById('pdfOverlay');
            if (overlay) overlay.style.display = 'none';
        } catch (_) {}
    }

    updatePdfProgress(current, total) {
        try {
            const text = document.getElementById('pdfOverlayText');
            const bar = document.getElementById('pdfOverlayBar');
            if (text) text.textContent = `학급 PDF 생성 중... (${current}/${total})`;
            if (bar) {
                const pct = Math.max(0, Math.min(100, Math.round((current / Math.max(1,total)) * 100)));
                bar.style.width = pct + '%';
            }
        } catch (_) {}
    }

    renderSubjectCards(student) {
        // 과목을 교과군별로 그룹화
        const groupedSubjects = {};
        const groupOrder = this.subjectGroups?.groups || {};

        this.combinedData.subjects.forEach(subject => {
            if (!this.hasStudentSubjectData(student, subject.name)) {
                return;
            }

            const groupName = this.getSubjectGroup(subject.name);
            if (!groupedSubjects[groupName]) {
                groupedSubjects[groupName] = {
                    subjects: [],
                    order: groupOrder[groupName]?.order || 99,
                    color: groupOrder[groupName]?.color || '#95a5a6'
                };
            }
            groupedSubjects[groupName].subjects.push(subject);
        });

        // 교과군 순서대로 정렬
        const sortedGroups = Object.entries(groupedSubjects)
            .sort((a, b) => a[1].order - b[1].order);

        if (sortedGroups.length === 0) {
            return '<p>표시할 과목 데이터가 없습니다.</p>';
        }

        // 교과군별로 테이블 생성
        return sortedGroups.map(([groupName, groupData]) => {
            const subjectRows = groupData.subjects.map(subject => {
                return this.renderSubjectTableRow(student, subject);
            }).join('');

            return `
                <div class="subject-group-section compact">
                    <div class="subject-group-header" style="border-left: 4px solid ${groupData.color}">
                        <h5>${groupName}</h5>
                        <span class="subject-count">${groupData.subjects.length}과목</span>
                    </div>
                    <table class="subject-table">
                        <thead>
                            <tr>
                                <th>과목</th>
                                <th>학점</th>
                                <th>점수</th>
                                <th>성취도</th>
                                <th>등급</th>
                                <th>석차</th>
                                <th>백분위</th>
                                <th>9등급</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${subjectRows}
                        </tbody>
                    </table>
                </div>
            `;
        }).join('');
    }

    hasStudentSubjectData(student, subjectName) {
        const hasOwn = (obj) => obj && Object.prototype.hasOwnProperty.call(obj, subjectName);
        return hasOwn(student.scores) ||
            hasOwn(student.achievements) ||
            hasOwn(student.grades) ||
            hasOwn(student.ranks) ||
            hasOwn(student.subjectTotals) ||
            hasOwn(student.percentiles);
    }

    // 개별 과목 테이블 행 렌더링
    renderSubjectTableRow(student, subject) {
        const hasScore = student.scores && Object.prototype.hasOwnProperty.call(student.scores, subject.name);
        const score = hasScore ? student.scores[subject.name] : null;
        const achievement = student.achievements && Object.prototype.hasOwnProperty.call(student.achievements, subject.name)
            ? student.achievements[subject.name]
            : '-';
        const grade = student.grades && Object.prototype.hasOwnProperty.call(student.grades, subject.name)
            ? student.grades[subject.name]
            : undefined;
        const rank = student.ranks && Object.prototype.hasOwnProperty.call(student.ranks, subject.name)
            ? student.ranks[subject.name]
            : '-';
        const percentile = student.percentiles && Object.prototype.hasOwnProperty.call(student.percentiles, subject.name)
            ? student.percentiles[subject.name]
            : null;

        const hasGrade = grade !== undefined && grade !== null && grade !== 'N/A' && !isNaN(grade);

        let percentileClass = 'low';
        if (hasGrade && percentile !== null && percentile >= 80) percentileClass = 'excellent';
        else if (hasGrade && percentile !== null && percentile >= 60) percentileClass = 'good';
        else if (hasGrade && percentile !== null && percentile >= 40) percentileClass = 'average';

        const grade9 = percentile !== null ? this.convertPercentileTo9Grade(percentile) : null;

        return `
            <tr class="${hasGrade ? '' : 'no-grade-row'}">
                <td class="subject-name-cell">${subject.name}</td>
                <td class="center">${subject.credits}</td>
                <td class="center">
                    <span class="score-value">${score !== null && score !== undefined ? score : '-'}</span>
                    <span class="avg-value">(${subject.average ? subject.average.toFixed(1) : '-'})</span>
                </td>
                <td class="center"><span class="achievement-badge ${achievement}">${achievement}</span></td>
                <td class="center">${hasGrade ? grade : '-'}</td>
                <td class="center">${rank}</td>
                <td class="center"><span class="percentile ${percentileClass}">${percentile !== null ? percentile + '%' : '-'}</span></td>
                <td class="center"><span class="${grade9 ? 'grade9-value' : ''}">${grade9 || '-'}</span></td>
            </tr>
        `;
    }

    createStudentPercentileChart(student) {
        const ctx = document.getElementById('studentPercentileChart');
        if (!ctx) return;

        // 기존 차트 제거
        if (this.studentPercentileChart) {
            this.studentPercentileChart.destroy();
        }

        // 교과군별 평균 등급 계산
        const groupGrades = this.calculateGroupGrades(student);

        // 데이터가 없으면 차트 숨김
        if (Object.keys(groupGrades).length === 0) {
            ctx.parentElement.style.display = 'none';
            return;
        }

        ctx.parentElement.style.display = 'block';

        // order 순으로 정렬
        const sortedGroups = Object.entries(groupGrades)
            .sort((a, b) => a[1].order - b[1].order);

        const labels = sortedGroups.map(([name, data]) => name);
        const gradeData = sortedGroups.map(([name, data]) => {
            // 등급을 역순으로 변환 (1등급=5, 2등급=4, ..., 5등급=1)하여 차트에서 높게 표시
            return 6 - data.averageGrade;
        });
        const colors = sortedGroups.map(([name, data]) => data.color);
        const originalGrades = sortedGroups.map(([name, data]) => data.averageGrade);
        const subjectDetails = sortedGroups.map(([name, data]) => data.subjects);

        this.studentPercentileChart = new Chart(ctx, {
            type: 'radar',
            plugins: [ChartDataLabels],
            data: {
                labels: labels,
                datasets: [{
                    label: '교과군별 평균등급',
                    data: gradeData,
                    backgroundColor: 'rgba(52, 152, 219, 0.2)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 2,
                    pointBackgroundColor: colors,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    pointRadius: 8
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
                            title: function(context) {
                                return context[0].label + ' 교과(군)';
                            },
                            label: function(context) {
                                const idx = context.dataIndex;
                                const avgGrade = originalGrades[idx];
                                return `평균 등급: ${avgGrade.toFixed(2)}등급`;
                            },
                            afterLabel: function(context) {
                                const idx = context.dataIndex;
                                const subjects = subjectDetails[idx];
                                if (subjects && subjects.length > 0) {
                                    const lines = ['포함 과목:'];
                                    subjects.forEach(s => {
                                        lines.push(`  ${s.name}: ${s.grade}등급 (${s.credits}학점)`);
                                    });
                                    return lines;
                                }
                                return '';
                            }
                        }
                    },
                    datalabels: {
                        display: true,
                        color: '#2c3e50',
                        backgroundColor: 'rgba(255, 255, 255, 0.9)',
                        borderColor: function(context) {
                            return colors[context.dataIndex];
                        },
                        borderWidth: 2,
                        borderRadius: 6,
                        padding: {
                            top: 6,
                            bottom: 6,
                            left: 8,
                            right: 8
                        },
                        font: {
                            size: 12,
                            weight: 'bold'
                        },
                        formatter: function(value, context) {
                            const idx = context.dataIndex;
                            return originalGrades[idx].toFixed(2) + '등급';
                        },
                        anchor: 'end',
                        align: 'top',
                        offset: 12,
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
                                size: 14,
                                weight: '600'
                            },
                            color: function(context) {
                                return colors[context.index] || '#2c3e50';
                            }
                        }
                    }
                }
            }
        });
    }

    // 프린터 출력 기능은 비활성화되었습니다.

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

    exportToCSV() {
        if (!this.combinedData || !this.combinedData.students || this.combinedData.students.length === 0) {
            this.showError('분석된 학생 데이터가 없습니다. 먼저 분석을 진행해주세요.');
            return;
        }

        // 개인정보 포함 여부를 묻는 모달 표시
        this.showExportOptionsModal();
    }

    showExportOptionsModal() {
        // 기존 모달이 있으면 제거
        const existingModal = document.getElementById('exportOptionsModal');
        if (existingModal) existingModal.remove();

        const modal = document.createElement('div');
        modal.id = 'exportOptionsModal';
        modal.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.5); display: flex; align-items: center;
            justify-content: center; z-index: 10000;
        `;
        modal.innerHTML = `
            <div style="background: var(--bg-card, #fff); padding: 30px; border-radius: 16px;
                        max-width: 400px; width: 90%; box-shadow: 0 10px 40px rgba(0,0,0,0.2);">
                <h3 style="margin: 0 0 20px 0; color: var(--text-primary, #333); font-size: 1.2rem;">
                    취합용 DB 파일 생성
                </h3>
                <div style="margin-bottom: 20px;">
                    <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;
                                  padding: 12px; background: var(--neutral-100, #f5f5f5);
                                  border-radius: 8px; user-select: none;">
                        <input type="checkbox" id="removePersonalInfo" checked
                               style="width: 18px; height: 18px; cursor: pointer;">
                        <span style="color: var(--text-primary, #333);">
                            개인정보 제외 (학번, 이름 삭제)
                        </span>
                    </label>
                    <p style="margin: 10px 0 0 0; font-size: 0.85rem; color: var(--text-muted, #888);">
                        체크 해제 시 A열: 학번, B열: 이름이 포함됩니다.
                    </p>
                </div>
                <div style="display: flex; gap: 10px; justify-content: flex-end;">
                    <button id="cancelExport" style="padding: 10px 20px; border: 1px solid var(--neutral-300, #ddd);
                            background: var(--bg-card, #fff); border-radius: 8px; cursor: pointer;
                            color: var(--text-secondary, #666);">
                        취소
                    </button>
                    <button id="confirmExport" style="padding: 10px 20px; border: none;
                            background: linear-gradient(135deg, var(--primary, #8B2942), var(--primary-dark, #6B1D32));
                            color: white; border-radius: 8px; cursor: pointer; font-weight: 500;">
                        CSV 다운로드
                    </button>
                </div>
            </div>
        `;

        document.body.appendChild(modal);

        // 이벤트 리스너
        document.getElementById('cancelExport').addEventListener('click', () => modal.remove());
        document.getElementById('confirmExport').addEventListener('click', () => {
            const removePersonalInfo = document.getElementById('removePersonalInfo').checked;
            modal.remove();
            this.generateCSV(removePersonalInfo);
        });

        // 모달 바깥 클릭 시 닫기
        modal.addEventListener('click', (e) => {
            if (e.target === modal) modal.remove();
        });
    }

    generateCSV(removePersonalInfo) {
        try {
            const subjects = this.combinedData.subjects;
            const groupOrder = this.subjectGroups?.groups || {};

            // 교과군 목록 추출 (순서대로)
            const subjectGroups = {};
            subjects.forEach(subject => {
                const groupName = this.getSubjectGroup(subject.name);
                if (!subjectGroups[groupName]) {
                    subjectGroups[groupName] = {
                        subjects: [],
                        order: groupOrder[groupName]?.order || 99
                    };
                }
                subjectGroups[groupName].subjects.push(subject);
            });
            const sortedGroupNames = Object.entries(subjectGroups)
                .sort((a, b) => a[1].order - b[1].order)
                .map(([name]) => name);

            // CSV 헤더 생성
            const headers = [];

            // 개인정보 열 (옵션)
            if (!removePersonalInfo) {
                headers.push('학번', '이름');
            }

            // 평균등급
            headers.push('평균등급(5등급)', '평균등급(9등급)');

            // 과목별 등급 (5등급)
            subjects.forEach(subject => {
                headers.push(this.getSubjectColumnLabel(subject));
            });

            // 교과군별 평균등급
            sortedGroupNames.forEach(groupName => {
                headers.push(`[${groupName}]평균`);
            });

            // 학생 데이터 정렬 (평균등급 오름차순)
            const sortedStudents = [...this.combinedData.students].sort((a, b) => {
                const gradeA = a.weightedAverageGrade || 999;
                const gradeB = b.weightedAverageGrade || 999;
                return gradeA - gradeB;
            });

            // CSV 데이터 생성
            const csvData = [headers];

            sortedStudents.forEach(student => {
                const row = [];

                // 개인정보 (옵션)
                if (!removePersonalInfo) {
                    const studentId = `${student.grade}${String(student.class).padStart(2, '0')}${String(student.originalNumber || student.number).padStart(2, '0')}`;
                    row.push(studentId, student.name || '');
                }

                // 평균등급 (5등급, 9등급)
                row.push(
                    student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : '',
                    student.weightedAverage9Grade ? student.weightedAverage9Grade.toFixed(2) : ''
                );

                // 과목별 등급 (5등급)
                subjects.forEach(subject => {
                    const grade = student.grades ? student.grades[subject.name] : '';
                    row.push(grade || '');
                });

                // 교과군별 평균등급 계산
                sortedGroupNames.forEach(groupName => {
                    const groupSubjects = subjectGroups[groupName].subjects;
                    let totalGrade = 0;
                    let totalCredits = 0;

                    groupSubjects.forEach(subject => {
                        const grade = student.grades ? student.grades[subject.name] : null;
                        if (grade && !isNaN(grade)) {
                            totalGrade += grade * (subject.credits || 1);
                            totalCredits += (subject.credits || 1);
                        }
                    });

                    const avgGrade = totalCredits > 0 ? (totalGrade / totalCredits).toFixed(2) : '';
                    row.push(avgGrade);
                });

                csvData.push(row);
            });

            // CSV 문자열로 변환
            const csvContent = csvData.map(row =>
                row.map(field => {
                    if (typeof field === 'string' && (field.includes(',') || field.includes('"') || field.includes('\n'))) {
                        return '"' + field.replace(/"/g, '""') + '"';
                    }
                    return field;
                }).join(',')
            ).join('\n');

            // BOM을 추가하여 한글이 제대로 표시되도록 함
            const BOM = '\uFEFF';
            const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });

            // 파일 다운로드
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);

            const now = new Date();
            const dateStr = now.getFullYear() +
                           String(now.getMonth() + 1).padStart(2, '0') +
                           String(now.getDate()).padStart(2, '0') + '_' +
                           String(now.getHours()).padStart(2, '0') +
                           String(now.getMinutes()).padStart(2, '0');

            link.setAttribute('download', `학생성적_취합DB_${dateStr}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);

            console.log(`CSV 파일이 생성되었습니다. 총 ${this.combinedData.students.length}명의 학생 데이터가 포함됩니다.`);

        } catch (error) {
            this.showError('CSV 파일 생성 중 오류가 발생했습니다: ' + error.message);
            console.error('CSV export error:', error);
        }
    }

    // 5등급을 9등급으로 환산하는 메소드
    convertTo9Grade(grade5) {
        if (!grade5 || grade5 < 1 || grade5 > 5) return '';
        
        // 5등급 → 9등급 환산표
        const conversionTable = {
            1: [1, 2],      // 1등급 → 1,2등급
            2: [3, 4],      // 2등급 → 3,4등급  
            3: [5, 6],      // 3등급 → 5,6등급
            4: [7, 8],      // 4등급 → 7,8등급
            5: [9]          // 5등급 → 9등급
        };
        
        const range = conversionTable[grade5];
        if (!range) return '';
        
        // 범위의 중간값 반환 (예: [1,2] → 1.5, [9] → 9)
        if (range.length === 1) {
            return range[0];
        } else {
            return (range[0] + range[1]) / 2;
        }
    }

    // 독립형 HTML 파일로 내보내기
    showHtmlExportOptionsModal() {
        const existingModal = document.getElementById('htmlExportOptionsModal');
        if (existingModal) existingModal.remove();

        const modal = document.createElement('div');
        modal.id = 'htmlExportOptionsModal';
        modal.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.5); display: flex; align-items: center;
            justify-content: center; z-index: 10000;
        `;
        modal.innerHTML = `
            <div style="background: var(--bg-card, #fff); padding: 30px; border-radius: 16px;
                        max-width: 440px; width: 92%; box-shadow: 0 10px 40px rgba(0,0,0,0.2);">
                <h3 style="margin: 0 0 18px 0; color: var(--text-primary, #333); font-size: 1.2rem;">
                    분석결과 HTML 저장
                </h3>
                <p style="margin: 0 0 14px 0; color: var(--text-secondary, #666); line-height: 1.5; font-size: 0.92rem;">
                    열기 암호를 설정하면 저장 파일을 열 때 먼저 암호를 입력해야 합니다.
                    비워두면 기존처럼 암호 없이 저장됩니다.
                </p>
                <div style="display: grid; gap: 12px; margin-bottom: 10px;">
                    <label style="display: grid; gap: 6px;">
                        <span style="font-size: 0.9rem; color: var(--text-primary, #333); font-weight: 600;">열기 암호</span>
                        <input type="password" id="htmlExportPassword" placeholder="선택 입력"
                               style="padding: 12px; border: 1px solid var(--neutral-300, #ddd); border-radius: 8px; font-size: 0.95rem;">
                    </label>
                    <label style="display: grid; gap: 6px;">
                        <span style="font-size: 0.9rem; color: var(--text-primary, #333); font-weight: 600;">암호 확인</span>
                        <input type="password" id="htmlExportPasswordConfirm" placeholder="암호를 다시 입력"
                               style="padding: 12px; border: 1px solid var(--neutral-300, #ddd); border-radius: 8px; font-size: 0.95rem;">
                    </label>
                </div>
                <p style="margin: 0 0 18px 0; font-size: 0.82rem; color: var(--text-muted, #888); line-height: 1.5;">
                    이 기능은 최소한의 보호 장치입니다. 암호를 잊으면 저장 파일에서 데이터를 복구할 수 없습니다.
                </p>
                <div id="htmlExportOptionsError" style="min-height: 1.2em; margin-bottom: 12px; color: #c0392b; font-size: 0.85rem;"></div>
                <div style="display: flex; gap: 10px; justify-content: flex-end;">
                    <button id="cancelHtmlExport" style="padding: 10px 20px; border: 1px solid var(--neutral-300, #ddd);
                            background: var(--bg-card, #fff); border-radius: 8px; cursor: pointer;
                            color: var(--text-secondary, #666);">
                        취소
                    </button>
                    <button id="confirmHtmlExport" style="padding: 10px 20px; border: none;
                            background: linear-gradient(135deg, var(--primary, #8B2942), var(--primary-dark, #6B1D32));
                            color: white; border-radius: 8px; cursor: pointer; font-weight: 500;">
                        HTML 다운로드
                    </button>
                </div>
            </div>
        `;

        document.body.appendChild(modal);

        const passwordInput = document.getElementById('htmlExportPassword');
        const confirmInput = document.getElementById('htmlExportPasswordConfirm');
        const errorDiv = document.getElementById('htmlExportOptionsError');
        const closeModal = () => modal.remove();

        document.getElementById('cancelHtmlExport').addEventListener('click', closeModal);
        document.getElementById('confirmHtmlExport').addEventListener('click', async () => {
            const password = passwordInput ? passwordInput.value : '';
            const passwordConfirm = confirmInput ? confirmInput.value : '';

            if ((password || passwordConfirm) && password !== passwordConfirm) {
                if (errorDiv) errorDiv.textContent = '암호와 확인 입력이 일치하지 않습니다.';
                if (confirmInput) confirmInput.focus();
                return;
            }

            closeModal();
            await this.exportAsStandaloneHtml({ password });
        });

        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeModal();
        });

        if (passwordInput) {
            setTimeout(() => passwordInput.focus(), 0);
        }
    }

    async exportAsStandaloneHtml(options = {}) {
        if (!this.combinedData) {
            this.showError('분석 데이터가 없습니다.');
            return;
        }

        try {
            // 현재 페이지의 HTML을 읽어서 독립형 버전 생성
            const htmlTemplate = await this.generateStandaloneHtmlTemplate(options);
            
            // BOM을 추가하여 한글이 제대로 표시되도록 함
            const BOM = '\uFEFF';
            const blob = new Blob([BOM + htmlTemplate], { type: 'text/html;charset=utf-8;' });

            // 파일 다운로드
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            
            // 파일명 생성 (현재 날짜 포함)
            const now = new Date();
            const dateStr = now.getFullYear() + 
                           String(now.getMonth() + 1).padStart(2, '0') + 
                           String(now.getDate()).padStart(2, '0') + '_' +
                           String(now.getHours()).padStart(2, '0') + 
                           String(now.getMinutes()).padStart(2, '0');
            
            link.setAttribute('download', `index_${dateStr}.html`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);

            console.log('독립형 HTML 파일이 생성되었습니다.');
            
        } catch (error) {
            this.showError('HTML 파일 생성 중 오류가 발생했습니다: ' + error.message);
            console.error('HTML export error:', error);
        }
    }

    getRuntimeScriptText() {
        try {
            if (typeof ScoreAnalyzer === 'function') {
                return `${ScoreAnalyzer.toString()}

let scoreAnalyzer;

document.addEventListener('DOMContentLoaded', () => {
    scoreAnalyzer = new ScoreAnalyzer();
});
`;
            }
        } catch (error) {
            console.warn('실행 중인 ScoreAnalyzer 소스 추출 실패:', error);
        }

        return '';
    }

    escapeInlineScriptContent(text) {
        if (!text) return '';
        return String(text).replace(/<\/script/gi, '<\\/script');
    }

    arrayBufferToBase64(buffer) {
        const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
        let binary = '';
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
        }
        return btoa(binary);
    }

    async derivePasswordKey(password, salt, keyUsages) {
        if (!window.crypto || !window.crypto.subtle) {
            throw new Error('현재 브라우저는 암호 보호 HTML 저장을 지원하지 않습니다.');
        }

        const encoder = new TextEncoder();
        const baseKey = await crypto.subtle.importKey(
            'raw',
            encoder.encode(password),
            'PBKDF2',
            false,
            ['deriveKey']
        );

        return crypto.subtle.deriveKey(
            {
                name: 'PBKDF2',
                salt,
                iterations: 250000,
                hash: 'SHA-256'
            },
            baseKey,
            {
                name: 'AES-GCM',
                length: 256
            },
            false,
            keyUsages
        );
    }

    async encryptExportPayload(password, payload) {
        const encoder = new TextEncoder();
        const salt = crypto.getRandomValues(new Uint8Array(16));
        const iv = crypto.getRandomValues(new Uint8Array(12));
        const key = await this.derivePasswordKey(password, salt, ['encrypt']);
        const ciphertext = await crypto.subtle.encrypt(
            {
                name: 'AES-GCM',
                iv
            },
            key,
            encoder.encode(JSON.stringify(payload))
        );

        return {
            salt: this.arrayBufferToBase64(salt),
            iv: this.arrayBufferToBase64(iv),
            ciphertext: this.arrayBufferToBase64(ciphertext),
            iterations: 250000
        };
    }

    buildProtectedHtmlBootstrap(encryptedPayload) {
        return `
(() => {
    const encryptedPayload = ${JSON.stringify(encryptedPayload)};
    const LOCK_CLASS = 'protected-export-locked';
    window.APP_BUILD_UTC = new Date().toISOString();

    const style = document.createElement('style');
    style.textContent = \`
body.\${LOCK_CLASS} .container { visibility: hidden !important; }
#protectedExportOverlay {
    position: fixed;
    inset: 0;
    z-index: 20000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 24px;
    background: rgba(15, 23, 42, 0.55);
    backdrop-filter: blur(8px);
}
#protectedExportOverlay .overlay-card {
    width: min(100%, 420px);
    background: #ffffff;
    border-radius: 20px;
    padding: 28px;
    box-shadow: 0 24px 60px rgba(15, 23, 42, 0.22);
    border: 1px solid rgba(15, 23, 42, 0.08);
}
#protectedExportOverlay h2 {
    margin: 0 0 10px 0;
    font-size: 1.25rem;
    color: #0f172a;
}
#protectedExportOverlay p {
    margin: 0 0 16px 0;
    color: #475569;
    line-height: 1.55;
    font-size: 0.93rem;
}
#protectedExportOverlay input {
    width: 100%;
    padding: 12px 14px;
    border-radius: 10px;
    border: 1px solid #cbd5e1;
    font-size: 1rem;
    margin-bottom: 12px;
    box-sizing: border-box;
}
#protectedExportOverlay button {
    width: 100%;
    border: none;
    border-radius: 10px;
    padding: 12px 14px;
    background: #10A37F;
    color: #ffffff;
    font-size: 0.98rem;
    font-weight: 600;
    cursor: pointer;
}
#protectedExportOverlay button:disabled {
    opacity: 0.6;
    cursor: wait;
}
#protectedExportStatus {
    min-height: 1.2em;
    margin-top: 12px;
    color: #b91c1c;
    font-size: 0.88rem;
}
\`;
    document.head.appendChild(style);
    document.body.classList.add(LOCK_CLASS);

    const overlay = document.createElement('div');
    overlay.id = 'protectedExportOverlay';
    overlay.innerHTML = \`
        <div class="overlay-card">
            <h2>암호 보호된 분석 결과</h2>
            <p>이 파일은 암호가 맞아야 분석 결과를 복호화해서 보여줍니다.</p>
            <input id="protectedExportPassword" type="password" placeholder="열기 암호 입력" autocomplete="current-password">
            <button id="protectedExportUnlock">열기</button>
            <div id="protectedExportStatus"></div>
        </div>
    \`;
    document.body.appendChild(overlay);

    const passwordInput = document.getElementById('protectedExportPassword');
    const unlockButton = document.getElementById('protectedExportUnlock');
    const statusEl = document.getElementById('protectedExportStatus');

    const base64ToBytes = (base64) => {
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
        }
        return bytes;
    };

    const unlock = async () => {
        const password = passwordInput ? passwordInput.value : '';
        if (!password) {
            if (statusEl) statusEl.textContent = '암호를 입력하세요.';
            if (passwordInput) passwordInput.focus();
            return;
        }

        unlockButton.disabled = true;
        if (statusEl) statusEl.textContent = '암호 확인 중...';

        try {
            const salt = base64ToBytes(encryptedPayload.salt);
            const iv = base64ToBytes(encryptedPayload.iv);
            const ciphertext = base64ToBytes(encryptedPayload.ciphertext);
            const encoder = new TextEncoder();
            const baseKey = await crypto.subtle.importKey(
                'raw',
                encoder.encode(password),
                'PBKDF2',
                false,
                ['deriveKey']
            );
            const key = await crypto.subtle.deriveKey(
                {
                    name: 'PBKDF2',
                    salt,
                    iterations: encryptedPayload.iterations,
                    hash: 'SHA-256'
                },
                baseKey,
                {
                    name: 'AES-GCM',
                    length: 256
                },
                false,
                ['decrypt']
            );
            const decrypted = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, ciphertext);
            const payload = JSON.parse(new TextDecoder().decode(decrypted));

            window.PRELOADED_DATA = payload.analysisData;
            window.PRELOADED_SUBJECT_GROUPS = payload.subjectGroups;
            window.PRELOADED_UI_STATE = payload.uiState;

            document.body.classList.remove(LOCK_CLASS);
            overlay.remove();

            const initializeDecryptedView = async () => {
                if (typeof scoreAnalyzer !== 'undefined' && scoreAnalyzer && typeof scoreAnalyzer.initializePreloadedView === 'function') {
                    scoreAnalyzer.subjectGroups = payload.subjectGroups || scoreAnalyzer.subjectGroups;
                    scoreAnalyzer.subjectGroupsReady = Promise.resolve(scoreAnalyzer.subjectGroups);
                    await scoreAnalyzer.initializePreloadedView();
                } else if (typeof ScoreAnalyzer === 'function') {
                    scoreAnalyzer = new ScoreAnalyzer();
                }
            };

            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', () => {
                    initializeDecryptedView();
                }, { once: true });
            } else {
                await initializeDecryptedView();
            }
        } catch (error) {
            if (statusEl) statusEl.textContent = '암호가 올바르지 않거나 파일이 손상되었습니다.';
            if (passwordInput) {
                passwordInput.focus();
                passwordInput.select();
            }
        } finally {
            unlockButton.disabled = false;
        }
    };

    unlockButton.addEventListener('click', () => {
        unlock();
    });

    if (passwordInput) {
        passwordInput.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                unlock();
            }
        });
        setTimeout(() => passwordInput.focus(), 0);
    }
})();
`;
    }

    // 독립형 HTML 템플릿 생성
    async generateStandaloneHtmlTemplate(options = {}) {
        const analysisData = JSON.stringify(this.combinedData);
        const subjectGroupsData = JSON.stringify(this.subjectGroups || null);
        const uiState = JSON.stringify(this.getCurrentUiState());
        const password = typeof options.password === 'string' ? options.password : '';

        // HTML 저장 시 업로드 섹션 임시 숨김
        const uploadSection = document.querySelector('.upload-section');
        const wasHidden = uploadSection ? uploadSection.style.display : '';
        if (uploadSection) uploadSection.style.display = 'none';

        // 원본 index.html, style.css, script.js를 그대로 사용하여 완전 동일한 구조로 생성
        const fetchText = async (url) => {
            try {
                const res = await fetch(url, { cache: 'no-cache' });
                if (!res || !res.ok) throw new Error('HTTP ' + (res && res.status));
                return await res.text();
            } catch (e) {
                console.warn('리소스 로드 실패:', url, e);
                return '';
            }
        };

        const [indexHtml, jsText, xlsx, chart, datalabels, jszip, jspdf, html2canvas] = await Promise.all([
            fetchText('index.html'),
            fetchText('script.js'),
            fetchText('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'),
            fetchText('https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js'),
            fetchText('https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2'),
            fetchText('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'),
            fetchText('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'),
            fetchText('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js')
        ]);
        const cssText = await this.getStyleCSS();
        const runtimeJsText = (jsText && jsText.trim()) ? jsText : this.getRuntimeScriptText();

        // DOMParser로 원본 index.html을 파싱하여 안전하게 조작
        const htmlSource = document.documentElement?.outerHTML || indexHtml;
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlSource, 'text/html');

        // 1) style.css 링크 -> 인라인 <style>
        try {
            const link = doc.querySelector('link[href="style.css"]');
            if (link && cssText) {
                const styleEl = doc.createElement('style');
                styleEl.textContent = cssText;
                link.replaceWith(styleEl);
            }
        } catch (_) {}

        // 2) 외부 라이브러리 <script src=...> 인라인 치환
        const inlineMap = new Map([
            ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', xlsx],
            ['https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js', chart],
            ['https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2', datalabels],
            ['https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', jszip],
            ['https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', jspdf],
            ['https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js', html2canvas]
        ]);

        doc.querySelectorAll('script[src]').forEach((s) => {
            const srcAttr = s.getAttribute('src');
            if (inlineMap.has(srcAttr) && inlineMap.get(srcAttr)) {
                const inline = doc.createElement('script');
                inline.textContent = this.escapeInlineScriptContent(inlineMap.get(srcAttr));
                s.replaceWith(inline);
            }
        });

        // ── HTML 저장용: 업로드 섹션 숨기고, 분석결과는 표시 ──
        const exportUploadSection = doc.querySelector('.upload-section');
        if (exportUploadSection) {
            exportUploadSection.style.display = 'none';
        }
        // 분석결과가 보이도록 설정
        const exportResults = doc.getElementById('results');
        if (exportResults) {
            exportResults.style.display = 'block';
        }

        // 3) script.js 인라인 및 PRELOADED_DATA 주입
        try {
            const appScript = doc.querySelector('script[src="script.js"]');
            const preload = doc.createElement('script');
            let preloadScript = `window.APP_BUILD_UTC = new Date().toISOString();\nwindow.PRELOADED_DATA = ${analysisData};\nwindow.PRELOADED_SUBJECT_GROUPS = ${subjectGroupsData};\nwindow.PRELOADED_UI_STATE = ${uiState};\nwindow.PRELOADED_HIDE_UPLOAD = true;`;
            if (password) {
                const encryptedPayload = await this.encryptExportPayload(password, {
                    analysisData: this.combinedData,
                    subjectGroups: this.subjectGroups || null,
                    uiState: this.getCurrentUiState()
                });
                preloadScript = this.buildProtectedHtmlBootstrap(encryptedPayload);
            }
            preload.textContent = this.escapeInlineScriptContent(preloadScript);
            const inline = doc.createElement('script');
            if (!runtimeJsText || !runtimeJsText.trim()) {
                console.warn('동작 스크립트를 확보하지 못해 현재 화면 스냅샷으로 저장합니다.');
                return await this.generateExactSnapshotHtmlTemplate();
            }
            inline.textContent = this.escapeInlineScriptContent(runtimeJsText);
            if (appScript) {
                appScript.replaceWith(preload);
                preload.after(inline);
            } else {
                doc.body.appendChild(preload);
                doc.body.appendChild(inline);
            }
        } catch (_) {}
        // 업로드 섹션 복원
        if (uploadSection) uploadSection.style.display = wasHidden;

        return '<!DOCTYPE html>' + doc.documentElement.outerHTML;
    }

    // CSS 파일 내용 가져오기 (style.css 우선, 실패 시 CSSOM, 최종 내장 CSS)
    async getStyleCSS() {
        // 1) style.css 직접 읽기 시도 (가장 확실하게 동일 스타일 보장)
        try {
            const res = await fetch('style.css', { cache: 'no-cache' });
            if (res && res.ok) {
                const text = await res.text();
                // 올바른 테마인지 확인 (보라색 primary 색상)
                if (text && text.trim().length > 0 && text.includes('#5F4A8B')) return text;
                // 내용이 있어도 구버전 테마면 내장 CSS 사용
                if (text && text.trim().length > 0) {
                    console.warn('로드된 style.css가 현재 테마와 다릅니다. 내장 CSS를 사용합니다.');
                }
            }
        } catch (_) {
            // 무시하고 다음 방법 시도
        }

        // 2) CSSOM에서 style.css 규칙 추출 (일부 환경에서 보안 정책으로 실패할 수 있음)
        try {
            const styleSheets = document.styleSheets;
            let cssText = '';
            for (let i = 0; i < styleSheets.length; i++) {
                try {
                    const styleSheet = styleSheets[i];
                    if (styleSheet.href && styleSheet.href.includes('style.css')) {
                        const rules = styleSheet.cssRules || styleSheet.rules;
                        for (let j = 0; j < rules.length; j++) {
                            cssText += rules[j].cssText + '\n';
                        }
                    }
                } catch (_) {
                    // 접근 불가한 경우 넘어감
                    continue;
                }
            }
            if (cssText.trim() && cssText.includes('#5F4A8B')) return cssText;
        } catch (_) {
            // 넘어가서 내장 CSS 사용
        }

        // 3) 현재 페이지의 <style> 태그에서 직접 CSS 텍스트 추출
        //    (저장된 HTML 스냅샷 또는 인라인 CSS가 있을 때)
        try {
            const styleTags = document.querySelectorAll('style');
            let cssText = '';
            for (let i = 0; i < styleTags.length; i++) {
                const text = styleTags[i].textContent || '';
                if (text.trim().length > 100 && text.includes('#5F4A8B')) {
                    cssText += text + '\n';
                }
            }
            if (cssText.trim()) {
                console.log('style 태그에서 올바른 테마 CSS 추출 성공 (' + cssText.length + ' chars)');
                return cssText;
            }
        } catch (_) {
            // 넘어가서 내장 CSS 사용
        }

        // 4) 최종 Fallback: 내장 CSS (항상 최신 보라색 테마 보장)
        console.log('내장 CSS(최신 테마)를 사용합니다.');
        return this.getBuiltInCSS();
    }

    // 내장 CSS 스타일 (현재 style.css와 동기화됨)
    getBuiltInCSS() {
        return `/* ========================================
   Modern Clean Theme
   ======================================== */

:root {
    --primary: #5F4A8B;
    --primary-light: #7B62A8;
    --primary-dark: #483670;
    --primary-bg: rgba(95, 74, 139, 0.08);

    --accent: #7B62A8;
    --accent-light: #9578C4;
    --accent-muted: #E8E0F4;
    --accent-bg: rgba(123, 98, 168, 0.08);

    --neutral-50: #FCFCFD;
    --neutral-100: #F8F7FA;
    --neutral-200: #EEEDF2;
    --neutral-300: #DDDBE5;
    --neutral-400: #B8B4C4;
    --neutral-500: #7E7891;
    --neutral-600: #585269;
    --neutral-700: #36314A;
    --neutral-800: #1A1626;

    --success: #15803D;
    --success-light: #16A34A;
    --success-bg: rgba(21, 128, 61, 0.1);

    --info: #2563EB;
    --info-light: #3B82F6;
    --info-bg: rgba(37, 99, 235, 0.1);

    --warning: #D97706;
    --warning-light: #F59E0B;
    --warning-bg: rgba(217, 119, 6, 0.12);

    --bg-body: radial-gradient(circle at top, #FFFFFF 0%, #F6F5F9 42%, #EEEDF2 100%);
    --bg-card: #FFFFFF;
    --bg-card-hover: #FFFFFF;
    --bg-section: linear-gradient(180deg, #FFFFFF 0%, #F8F7FA 100%);

    --text-primary: #1A1626;
    --text-secondary: #585269;
    --text-muted: #7E7891;
    --text-inverse: #FFFFFF;

    --border-light: rgba(95, 74, 139, 0.08);
    --border-medium: rgba(95, 74, 139, 0.14);
    --border-accent: rgba(95, 74, 139, 0.18);

    --shadow-sm: 0 1px 2px rgba(95, 74, 139, 0.06);
    --shadow-md: 0 8px 24px rgba(95, 74, 139, 0.08);
    --shadow-lg: 0 16px 36px rgba(95, 74, 139, 0.10);
    --shadow-xl: 0 24px 64px rgba(95, 74, 139, 0.12);

    --radius-sm: 10px;
    --radius-md: 14px;
    --radius-lg: 20px;
    --radius-xl: 28px;

    --grade-1: #15803D;
    --grade-2: #5F4A8B;
    --grade-3: #0369A1;
    --grade-4: #D97706;
    --grade-5: #7E7891;
}

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: 'Pretendard Variable', 'Pretendard', 'SUIT Variable', 'Noto Sans KR', sans-serif;
    background: var(--bg-body);
    min-height: 100vh;
    padding: 18px;
    position: relative;
    overflow-x: hidden;
    color: var(--text-primary);
}

.container {
    max-width: 1320px;
    margin: 0 auto;
    background: var(--bg-card);
    border-radius: var(--radius-xl);
    box-shadow: var(--shadow-xl);
    overflow: hidden;
    position: relative;
    border: 1px solid var(--border-light);
}

header {
    background: linear-gradient(135deg, #5F4A8B 0%, #483670 50%, #36314A 100%);
    color: var(--text-inverse);
    padding: 40px 36px 36px;
    text-align: center;
    border-bottom: none;
    position: relative;
    overflow: hidden;
}

header::before {
    content: '✦';
    position: absolute;
    top: 16px;
    right: 24px;
    font-size: 1.5rem;
    opacity: 0.3;
    color: #FFFFFF;
}

header h1 {
    font-size: 2rem;
    margin-bottom: 6px;
    font-weight: 600;
    letter-spacing: -0.04em;
    position: relative;
    color: #FFFFFF;
}

.header-subtitle {
    font-size: 1rem;
    color: rgba(255, 255, 255, 0.85);
    max-width: 720px;
    line-height: 1.6;
    margin: 0 auto;
}

.badge-lite {
    display: inline-block;
    margin-left: 10px;
    padding: 4px 10px;
    font-size: 0.74rem;
    font-weight: 600;
    color: #FFFFFF;
    background: rgba(255, 255, 255, 0.2);
    border: 1px solid rgba(255, 255, 255, 0.4);
    border-radius: 999px;
    letter-spacing: 0.08em;
    vertical-align: middle;
    text-transform: uppercase;
}

.upload-section {
    padding: 32px 36px;
    text-align: center;
    border-bottom: 1px solid var(--neutral-200);
    position: relative;
    background: linear-gradient(180deg, var(--neutral-50) 0%, var(--bg-card) 100%);
}

.container.post-analysis .upload-guide,
.container.post-analysis .section-divider,
.container.post-analysis .file-input-wrapper,
.container.post-analysis #fileList,
.container.post-analysis #analyzeBtn {
    display: none !important;
}

.file-input-wrapper {
    margin-bottom: 30px;
}

.file-input-wrapper input[type="file"] {
    display: none;
}

.file-input-label {
    display: flex;
    align-items: center;
    justify-content: center;
    width: min(100%, 720px);
    min-height: 88px;
    margin: 0 auto;
    padding: 18px 28px;
    background: var(--bg-card);
    border: 1.5px dashed var(--neutral-300);
    border-radius: var(--radius-lg);
    cursor: pointer;
    transition: all 0.25s ease;
    font-size: 1rem;
    font-weight: 500;
    color: var(--text-primary);
}

.upload-hint {
    margin-top: 12px;
    font-size: 0.9rem;
    color: var(--text-muted);
}

.file-input-label:hover {
    background: var(--neutral-50);
    border-color: var(--primary);
    color: var(--primary-dark);
}

.file-input-label.dragover {
    background: var(--primary-bg);
    border-color: var(--primary);
    color: var(--primary);
    box-shadow: 0 0 0 4px var(--primary-bg) inset;
}

/* 업로드 섹션 전체 드래그오버 강조 및 오버레이 안내 */
.upload-section.dragover {
    border: 1.5px dashed var(--primary);
    border-radius: var(--radius-lg);
    background: var(--primary-bg);
}
.upload-section.dragover::after {
    content: '여기에 파일을 드롭하세요';
    position: absolute;
    left: 50%;
    top: 50%;
    transform: translate(-50%, -50%);
    color: var(--primary);
    font-weight: 600;
    font-size: 1rem;
    padding: 12px 18px;
    background: var(--bg-card);
    border: 1px solid var(--primary-light);
    border-radius: var(--radius-sm);
    pointer-events: none;
    box-shadow: var(--shadow-lg);
}

.analyze-btn {
    background: var(--primary);
    color: var(--text-inverse);
    border: 1px solid transparent;
    padding: 12px 22px;
    border-radius: 999px;
    font-size: 0.95rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.25s ease;
    box-shadow: none;
}

.analyze-btn:hover:not(:disabled) {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
    background: var(--primary-dark);
}

.analyze-btn:active:not(:disabled) {
    transform: translateY(0);
}

.analyze-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
    background: var(--neutral-400);
    box-shadow: none;
}

.action-buttons {
    display: flex;
    justify-content: center;
    gap: 10px;
    flex-wrap: wrap;
}

.secondary-btn {
    background: var(--bg-card);
    color: var(--text-primary);
    border-color: var(--neutral-300);
}

.secondary-btn:hover:not(:disabled) {
    background: var(--neutral-100);
    color: var(--text-primary);
    box-shadow: var(--shadow-sm);
}

.export-btn {
    background: var(--bg-card);
    color: var(--primary);
    border: 2px solid var(--primary);
    padding: 13px 32px;
    border-radius: var(--radius-lg);
    font-size: 1rem;
    cursor: pointer;
    transition: all 0.25s ease;
}

.export-btn:hover:not(:disabled) {
    background: var(--primary);
    color: var(--text-inverse);
    transform: translateY(-2px);
    box-shadow: var(--shadow-lg);
}

.export-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}

.upload-guide {
    background: var(--bg-card);
    border: 1px solid var(--neutral-200);
    padding: 22px 24px;
    margin: 0 auto 18px;
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow-sm);
    max-width: 920px;
    text-align: left;
}

.upload-guide p {
    margin: 0;
    color: var(--text-secondary);
    line-height: 1.6;
}

.guide-title {
    color: var(--text-primary);
    margin-bottom: 8px !important;
    font-size: 1rem;
    font-weight: 700;
}

.upload-guide strong {
    color: var(--text-primary);
}

.section-divider {
    height: 1px;
    background: linear-gradient(90deg, rgba(0,0,0,0.06), rgba(0,0,0,0.12), rgba(0,0,0,0.06));
    margin: 16px 0 22px 0;
    border: none;
}

.warning-text {
    color: var(--warning);
    font-weight: 600;
    font-size: 0.95rem;
    margin-top: 10px;
    text-align: left;
    padding: 8px 12px;
    background-color: var(--warning-bg);
    border-radius: var(--radius-md);
    border-left: 4px solid var(--warning);
}

/* 강조 색상: XLS vs XLS data 구분 표시 */
.warning-text .xls {
    color: var(--warning);
    background: rgba(217, 119, 6, 0.12);
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 800;
}
.warning-text .xlsdata {
    color: var(--success);
    background: var(--success-bg);
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 800;
}

.privacy-notice {
    margin-top: 10px;
    padding: 14px 16px;
    border-radius: var(--radius-md);
    background: var(--neutral-100);
    border: 1px solid var(--neutral-200);
    color: var(--text-secondary);
}
.privacy-notice p {
    margin: 0 0 8px 0;
    font-weight: 600;
    color: var(--text-primary);
}
.privacy-notice ul {
    margin: 0;
    padding-left: 18px;
    list-style: disc;
    color: var(--text-secondary);
}
.privacy-notice li {
    margin: 3px 0;
    line-height: 1.5;
}
.privacy-notice .privacy-footnote {
    color: var(--text-muted);
    opacity: 1;
    margin-top: 8px;
}

.results-section {
    padding: 28px 36px 36px;
}

/* 하단 크레딧 푸터 */
.app-footer {
    padding: 16px 36px 24px 36px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    border-top: 1px solid var(--neutral-200);
    background: none;
}
.app-footer .footer-right {
    display: flex;
    align-items: center;
    gap: 12px;
}
.app-footer .credits {
    text-align: right;
    font-size: 0.85rem;
    color: var(--text-secondary);
    background: none;
    padding: 0;
    border-radius: 0;
}
.app-footer .credits a:not(.help-btn) {
    color: var(--text-muted);
    text-decoration: none;
    border-bottom: 1px dashed var(--neutral-400);
}
.app-footer .credits a:not(.help-btn):hover {
    color: var(--text-primary);
    border-bottom-color: var(--neutral-500);
}

/* last updated 표시 */
.app-footer .updated {
    font-size: 0.8rem;
    color: var(--text-muted);
    margin-left: 8px;
}

/* 도움말 버튼 */
.help-btn {
    display: inline-block;
    padding: 6px 12px;
    font-size: 0.85rem;
    line-height: 1;
    border-radius: 999px;
    color: var(--primary);
    background: var(--bg-card);
    border: 1px solid var(--primary);
    text-decoration: none;
    transition: all 0.2s ease;
}
.help-btn:hover {
    color: var(--text-inverse);
    background: var(--primary);
    border-color: var(--primary);
    box-shadow: var(--shadow-sm);
}

.tabs {
    display: flex;
    gap: 0;
    padding: 0;
    background: none;
    border: none;
    border-bottom: 2px solid var(--neutral-200);
    border-radius: 0;
    margin-bottom: 30px;
}

.tab-btn {
    flex: 1;
    min-width: 0;
    padding: 14px 18px;
    background: none;
    border: none;
    border-bottom: 3px solid transparent;
    cursor: pointer;
    font-size: 0.95rem;
    font-weight: 500;
    color: var(--text-secondary);
    transition: all 0.25s ease;
    border-radius: 0;
    position: relative;
    margin-bottom: -2px;
}

.tab-btn.active {
    color: var(--primary);
    background: none;
    box-shadow: none;
    border-bottom-color: var(--primary);
    font-weight: 700;
}

.tab-btn:hover:not(.active) {
    background: none;
    color: var(--text-primary);
    border-bottom-color: var(--neutral-300);
}

.tab-content {
    display: none;
}

.tab-content.active {
    display: block;
}

.tab-content h2 {
    color: var(--text-primary);
    margin-bottom: 22px;
    font-size: 1.45rem;
    font-weight: 600;
}


.subject-averages {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 20px;
}

.subject-item {
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-50) 100%);
    border-radius: var(--radius-lg);
    padding: 22px;
    border: 1px solid var(--neutral-200);
    transition: all 0.25s ease;
    box-shadow: none;
}

.subject-item:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.subject-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
}

.subject-header h3 {
    color: var(--text-primary);
    font-size: 1.15rem;
    font-weight: 600;
}

.credits {
    background: var(--neutral-100);
    color: var(--text-secondary);
    padding: 5px 10px;
    border-radius: 15px;
    border: 1px solid var(--neutral-200);
    font-size: 0.8rem;
    font-weight: 500;
}

.average-score {
    text-align: center;
}

.average-score .score {
    display: block;
    font-size: 2.2rem;
    font-weight: 600;
    color: var(--text-primary);
    margin-bottom: 5px;
}

.average-score .label {
    color: var(--text-secondary);
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.achievement-bars {
    margin-top: 20px;
    padding-top: 20px;
    border-top: 1px solid var(--neutral-200);
}

.achievement-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 8px;
}

.achievement-label {
    width: 25px;
    font-weight: 600;
    color: var(--text-primary);
    font-size: 0.9rem;
    text-align: center;
}

.achievement-bar-container {
    flex: 1;
    height: 20px;
    background: var(--neutral-200);
    border-radius: 10px;
    overflow: hidden;
    position: relative;
}

.achievement-bar-fill {
    height: 100%;
    border-radius: 10px;
    transition: width 0.8s ease;
    min-width: 2px;
}

.achievement-bar:nth-child(1) .achievement-bar-fill { background: linear-gradient(135deg, var(--success), var(--success-light)); }
.achievement-bar:nth-child(2) .achievement-bar-fill { background: linear-gradient(135deg, var(--info), var(--info-light)); }
.achievement-bar:nth-child(3) .achievement-bar-fill { background: linear-gradient(135deg, var(--accent), var(--accent-light)); }
.achievement-bar:nth-child(4) .achievement-bar-fill { background: linear-gradient(135deg, var(--warning), var(--primary-light)); }
.achievement-bar:nth-child(5) .achievement-bar-fill { background: linear-gradient(135deg, var(--primary), var(--primary-dark)); }

.achievement-percentage {
    width: 50px;
    text-align: right;
    font-weight: 500;
    color: var(--text-primary);
    font-size: 0.85rem;
}

.achievement-distribution {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 30px;
}

.distribution-item {
    background: var(--bg-card);
    border-radius: var(--radius-md);
    padding: 25px;
    box-shadow: var(--shadow-sm);
    transition: all 0.25s ease;
}

.distribution-item:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.distribution-item h3 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.3rem;
    font-weight: 500;
}

.distribution-bars {
    display: flex;
    flex-direction: column;
    gap: 15px;
}

.grade-bar {
    display: flex;
    align-items: center;
    gap: 15px;
}

.grade-label {
    width: 30px;
    font-weight: 600;
    color: var(--text-primary);
    font-size: 1.1rem;
}

.bar-container {
    flex: 1;
    height: 25px;
    background: var(--neutral-200);
    border-radius: 15px;
    overflow: hidden;
    position: relative;
}

.bar {
    height: 100%;
    background: linear-gradient(135deg, var(--info) 0%, var(--info-light) 100%);
    border-radius: 15px;
    transition: width 0.8s ease;
    min-width: 2px;
}

.percentage {
    width: 60px;
    text-align: right;
    font-weight: 500;
    color: var(--text-primary);
}

.student-analysis {
    width: 100%;
}

.search-box {
    margin-bottom: 25px;
}

.search-box input {
    width: 100%;
    max-width: 400px;
    padding: 14px 18px;
    border: 1px solid var(--neutral-300);
    border-radius: 14px;
    font-size: 1rem;
    outline: none;
    transition: all 0.25s ease;
    background: var(--bg-card);
}

.search-box input:focus {
    border-color: var(--primary);
    box-shadow: 0 0 0 3px var(--primary-bg);
}

.students-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 18px;
    margin-top: 20px;
}

.student-card {
    background: var(--bg-card);
    border-radius: var(--radius-lg);
    box-shadow: none;
    border: 1px solid var(--neutral-200);
    transition: all 0.25s ease;
    overflow: hidden;
}

.student-card:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.student-card-header {
    background: linear-gradient(135deg, #5F4A8B 0%, #483670 100%);
    color: var(--text-inverse);
    padding: 16px 18px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    border-bottom: none;
}

.student-basic-info {
    display: flex;
    align-items: baseline;
    gap: 10px;
    flex-wrap: wrap;
}

.student-basic-info h4 {
    margin: 0;
    font-size: 1.2rem;
    font-weight: 600;
    white-space: nowrap;
}

.student-number {
    font-size: 0.88rem;
    color: var(--text-secondary);
    opacity: 1;
}

.student-summary {
    display: flex;
    flex-direction: row;
    gap: 6px;
    flex-wrap: wrap;
}

.summary-row {
    display: contents;
}

.summary-metric-inline {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    background: var(--neutral-100);
    border: 1px solid var(--neutral-200);
    padding: 4px 8px;
    border-radius: 6px;
    white-space: nowrap;
    flex-shrink: 0;
}

.summary-metric-inline .metric-label {
    font-size: 0.7rem;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    white-space: nowrap;
}

.summary-metric-inline .metric-value {
    font-size: 0.9rem;
    font-weight: 700;
    white-space: nowrap;
    color: var(--text-primary);
}

.summary-metric {
    text-align: center;
    background: rgba(255, 255, 255, 0.15);
    padding: 8px 12px;
    border-radius: 8px;
    min-width: 70px;
}

.summary-metric .metric-label {
    display: block;
    font-size: 0.7rem;
    opacity: 0.8;
    margin-bottom: 2px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.summary-metric .metric-value {
    display: block;
    font-size: 1.1rem;
    font-weight: 700;
}

.student-subjects {
    padding: 15px 20px;
    max-height: none;
    overflow: visible;
}

.subject-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 0;
    border-bottom: 1px solid var(--neutral-200);
}

.subject-row:last-child {
    border-bottom: none;
}

.subject-row.no-grade {
    opacity: 0.7;
}

.subject-name {
    font-weight: 500;
    color: var(--text-primary);
    flex: 1;
    font-size: 0.9rem;
}

.subject-data {
    display: flex;
    gap: 8px;
    align-items: center;
}

.subject-score {
    font-weight: 600;
    color: var(--primary);
    font-size: 0.85rem;
    min-width: 45px;
    text-align: right;
}

.subject-achievement {
    font-size: 0.8rem;
    padding: 2px 6px;
    border-radius: 3px;
    min-width: 20px;
    text-align: center;
}

.subject-grade {
    font-weight: 500;
    color: var(--text-secondary);
    font-size: 0.8rem;
    min-width: 35px;
    text-align: center;
}

.subject-percentile {
    font-weight: 500;
    color: var(--success);
    font-size: 0.8rem;
    min-width: 40px;
    text-align: right;
}

/* ── 학생 카드: 새 뱃지 레이아웃 ── */
.student-card-title-row {
    display: flex;
    align-items: baseline;
    gap: 10px;
    flex-wrap: wrap;
}

.student-card-name {
    margin: 0;
    font-size: 1.15rem;
    font-weight: 700;
    color: #FFFFFF;
    white-space: nowrap;
}

.student-card-class {
    font-size: 0.8rem;
    color: rgba(255, 255, 255, 0.75);
    white-space: nowrap;
}

.student-card-badges {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}

.card-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 5px 12px;
    background: rgba(255, 255, 255, 0.15);
    border: 1px solid rgba(255, 255, 255, 0.3);
    border-radius: 8px;
    white-space: nowrap;
    flex-shrink: 0;
    transition: border-color 0.2s ease;
}

.card-badge:hover {
    border-color: var(--neutral-300);
}

.card-badge-label {
    font-size: 0.68rem;
    font-weight: 600;
    color: rgba(255, 255, 255, 0.7);
    text-transform: uppercase;
    letter-spacing: 0.3px;
}

.card-badge-value {
    font-size: 0.88rem;
    font-weight: 700;
    color: #FFFFFF;
}

.card-badge-value.primary {
    color: #FFFFFF;
}

.card-badge-value.accent {
    color: #E8E0F4;
}

.card-badge-value small {
    font-size: 0.7rem;
    font-weight: 500;
    color: var(--text-muted);
}

.student-card-footer {
    background: var(--bg-card);
    padding: 15px 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid var(--neutral-200);
}

.grade-subjects-count {
    font-size: 0.85rem;
    color: var(--text-secondary);
}

.view-detail-btn {
    background: var(--bg-card);
    color: var(--text-primary);
    border: 1px solid var(--neutral-300);
    padding: 8px 16px;
    border-radius: var(--radius-sm);
    font-size: 0.85rem;
    cursor: pointer;
    transition: all 0.25s ease;
    font-weight: 500;
}

.view-detail-btn:hover {
    transform: translateY(-1px);
    border-color: var(--primary);
    color: var(--primary-dark);
    box-shadow: var(--shadow-sm);
}

.achievement.A {
    background: var(--success);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.B {
    background: var(--info);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.C {
    background: var(--accent);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.D {
    background: var(--warning);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.achievement.E, .achievement.미도달 {
    background: var(--primary);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.score {
    font-weight: 600;
    color: var(--text-primary);
}

.grade {
    text-align: center;
    font-weight: 500;
}

.rank {
    text-align: center;
    font-weight: 500;
    color: var(--text-secondary);
}

.avg-grade {
    text-align: center;
    font-weight: 600;
    color: var(--primary);
    font-size: 1.1rem;
}

.grade-analysis-container {
    display: grid;
    grid-template-columns: 1fr 1fr;
    grid-template-rows: auto auto;
    gap: 20px;
    margin-bottom: 30px;
}

.chart-section {
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-50) 100%);
    border-radius: var(--radius-lg);
    padding: 22px;
    text-align: center;
    border: 1px solid var(--neutral-200);
}

.chart-section h3 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.3rem;
    font-weight: 500;
}

.chart-section canvas {
    max-width: 100%;
    height: 350px !important;
}

.stats-section {
    grid-column: 1 / -1;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-50) 100%);
    border-radius: var(--radius-lg);
    padding: 22px;
    border: 1px solid var(--neutral-200);
}

.stat-item {
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    padding: 20px;
    box-shadow: none;
    border: 1px solid var(--neutral-200);
}

.stat-label {
    color: var(--text-secondary);
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 8px;
}

.stat-value {
    color: var(--text-primary);
    font-size: 2rem;
    font-weight: 600;
}

@media (max-width: 768px) {
    .grade-analysis-container {
        grid-template-columns: 1fr;
        gap: 20px;
    }
    
    .stats-section {
        grid-template-columns: repeat(2, 1fr);
        gap: 15px;
    }
    
    .chart-section {
        padding: 15px;
    }
    
    .stat-item {
        padding: 15px;
    }
    
    .stat-value {
        font-size: 1.5rem;
    }
}

.loading {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 60px;
    color: var(--text-secondary);
}

.spinner {
    width: 50px;
    height: 50px;
    border: 4px solid var(--neutral-200);
    border-top: 4px solid var(--primary);
    border-radius: 50%;
    animation: spin 1s linear infinite;
    margin-bottom: 20px;
}

@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}

.loading p {
    font-size: 1.1rem;
    color: var(--text-secondary);
}

.error-message {
    background: var(--warning-bg);
    color: var(--warning);
    padding: 20px;
    margin: 20px 40px;
    border-radius: var(--radius-md);
    border-left: 5px solid var(--warning);
    font-size: 1rem;
}

.file-list {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 20px;
    margin: 20px 0;
    border: 1px solid var(--neutral-200);
}

.file-list h4 {
    color: var(--text-primary);
    margin-bottom: 15px;
    font-size: 1.1rem;
}

.file-list ul {
    list-style: none;
    padding: 0;
}

.file-list li {
    background: var(--bg-card);
    padding: 10px 15px;
    margin: 8px 0;
    border-radius: var(--radius-sm);
    border: 1px solid var(--neutral-200);
    box-shadow: none;
}

.file-selector-section {
    background: var(--bg-card);
    padding: 20px;
    border-radius: var(--radius-lg);
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 15px;
    border: 1px solid var(--neutral-200);
}

.file-selector-section label {
    color: var(--text-primary);
    font-weight: 500;
    white-space: nowrap;
}

.file-select {
    flex: 1;
    padding: 10px 15px;
    border: 2px solid var(--neutral-300);
    border-radius: var(--radius-sm);
    font-size: 1rem;
    background: var(--bg-card);
    outline: none;
    transition: all 0.25s ease;
}

.file-select:focus {
    border-color: var(--primary);
    box-shadow: 0 0 0 3px var(--primary-bg);
}

.comparison-container {
    display: grid;
    grid-template-columns: 1fr;
    gap: 30px;
}

.comparison-section {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 25px;
    border: 1px solid var(--border-light);
}

.comparison-section h3 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.3rem;
    font-weight: 500;
}

.comparison-table {
    width: 100%;
    border-collapse: collapse;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    overflow: hidden;
    box-shadow: var(--shadow-sm);
}

.comparison-table th,
.comparison-table td {
    padding: 12px 15px;
    text-align: center;
    border-bottom: 1px solid var(--neutral-200);
}

.comparison-table th {
    background: linear-gradient(135deg, var(--primary) 0%, var(--primary-dark) 100%);
    color: var(--text-inverse);
    font-weight: 500;
    font-size: 0.9rem;
}

.comparison-table tr:nth-child(even) {
    background: var(--neutral-100);
}

.comparison-table tr:hover {
    background: var(--primary-bg);
}

@media (max-width: 768px) {
    .file-selector-section {
        flex-direction: column;
        align-items: stretch;
        gap: 10px;
    }
    
    .file-select {
        width: 100%;
    }
    
    .comparison-table {
        font-size: 0.8rem;
    }
    
    .comparison-table th,
    .comparison-table td {
        padding: 8px 6px;
    }
}

/* 학생 선택 및 상세 분석 스타일 */
.student-selector {
    display: flex;
    align-items: center;
    gap: 15px;
    background: var(--bg-card);
    padding: 20px;
    border-radius: var(--radius-lg);
    margin-bottom: 20px;
    flex-wrap: wrap;
    border: 1px solid var(--neutral-200);
    box-shadow: var(--shadow-sm);
}

.selector-group {
    display: flex;
    align-items: center;
    gap: 8px;
}

.selector-group label {
    font-weight: 500;
    color: var(--text-primary);
    white-space: nowrap;
}

.selector {
    padding: 10px 12px;
    border: 1px solid var(--neutral-300);
    border-radius: var(--radius-sm);
    font-size: 0.9rem;
    background: var(--bg-card);
    min-width: 120px;
    transition: all 0.25s ease;
}

.selector:focus {
    border-color: var(--primary);
    outline: none;
    box-shadow: 0 0 0 3px var(--primary-bg);
}

.detail-btn {
    background: var(--bg-card);
    color: var(--text-primary);
    border: 1px solid var(--neutral-300);
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    font-size: 0.9rem;
    cursor: pointer;
    transition: all 0.25s ease;
    white-space: nowrap;
}

.detail-btn:hover:not(:disabled) {
    transform: translateY(-2px);
    box-shadow: var(--shadow-sm);
    border-color: var(--primary);
    color: var(--primary-dark);
}

.detail-btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
}

.view-toggle {
    display: flex;
    background: var(--neutral-100);
    border-radius: var(--radius-md);
    padding: 4px;
    margin-bottom: 20px;
    width: fit-content;
    border: 1px solid var(--neutral-200);
}

.toggle-btn {
    background: none;
    border: none;
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    cursor: pointer;
    transition: all 0.25s ease;
    font-weight: 500;
    color: var(--text-secondary);
}

.toggle-btn.active {
    background: var(--bg-card);
    color: var(--text-primary);
    box-shadow: var(--shadow-sm);
}

/* 학생 상세 분석 스타일 */
.student-detail-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    background: linear-gradient(180deg, var(--bg-card) 0%, var(--neutral-100) 100%);
    color: var(--text-primary);
    padding: 22px 24px;
    border-radius: var(--radius-lg);
    margin-bottom: 20px;
    border: 1px solid var(--border-light);
    box-shadow: var(--shadow-md);
}

.student-info h3 {
    font-size: 1.5rem;
    margin-bottom: 6px;
    font-weight: 400;
}

.student-meta {
    display: flex;
    gap: 14px;
    font-size: 0.85rem;
    opacity: 0.9;
    flex-wrap: wrap;
}

.overall-stats {
    display: flex;
    gap: 12px;
}

.stat-card {
    text-align: center;
    background: var(--bg-card);
    padding: 12px 16px;
    border-radius: var(--radius-md);
    border: 1px solid var(--border-light);
    box-shadow: var(--shadow-sm);
    min-width: 110px;
}

.stat-label {
    display: block;
    font-size: 0.8rem;
    opacity: 0.8;
    margin-bottom: 5px;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.stat-value {
    display: block;
    font-size: 1.35rem;
    font-weight: 700;
    color: var(--text-primary);
}

.stat-value.grade {
    color: var(--primary);
}

.student-detail-content {
    display: flex;
    flex-direction: column;
    gap: 22px;
    margin-bottom: 24px;
}

.analysis-overview {
    display: grid;
    grid-template-columns: minmax(0, 1.35fr) minmax(280px, 0.85fr);
    gap: 20px;
    margin-bottom: 20px;
    align-items: start;
}

.student-summary {
    display: flex;
    flex-direction: column;
    gap: 20px;
}

.summary-card {
    background: var(--bg-card);
    border-radius: var(--radius-lg);
    padding: 18px 20px;
    box-shadow: var(--shadow-md);
    border: 1px solid var(--border-light);
}

.summary-header {
    margin-bottom: 14px;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--neutral-200);
}

.summary-header h4 {
    color: var(--text-primary);
    font-size: 1.2rem;
    font-weight: 600;
    margin: 0;
}

.summary-grid {
    display: grid;
    gap: 10px;
}

.summary-item {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 0;
    border-bottom: 1px solid var(--border-light);
}

.summary-item:last-child {
    border-bottom: none;
}

.summary-label {
    font-weight: 500;
    color: var(--text-secondary);
    font-size: 0.9rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.summary-value {
    font-weight: 600;
    color: var(--text-primary);
    font-size: 1rem;
    text-align: right;
}

.summary-value-group {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 4px;
}

.summary-value.highlight {
    color: var(--primary);
    font-size: 1.05rem;
    font-weight: 600;
}

.summary-value.orange {
    color: var(--accent);
    font-size: 1.05rem;
    font-weight: 600;
}

.summary-note {
    font-size: 0.72rem;
    color: var(--text-muted);
    text-align: right;
    line-height: 1.35;
    white-space: nowrap;
}

.metric-value.orange {
    color: var(--accent);
    font-weight: 600;
}

.chart-container {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 16px 18px 18px;
    text-align: center;
    border: 1px solid var(--border-light);
    width: 100%;
    max-width: 420px;
    justify-self: end;
}

.chart-container h4 {
    color: var(--text-primary);
    margin-bottom: 12px;
    font-size: 1rem;
    font-weight: 600;
}

.chart-container canvas {
    display: block;
    width: min(100%, 320px) !important;
    height: auto !important;
    margin: 0 auto;
}

.subject-details h4 {
    color: var(--text-primary);
    margin-bottom: 20px;
    font-size: 1.2rem;
    font-weight: 500;
}

.subject-cards {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 16px;
    max-height: none;
}

@media (max-width: 1200px) {
    .subject-cards {
        grid-template-columns: 1fr;
    }
}

/* 교과(군)별 섹션 스타일 */
.subject-group-section {
    background: var(--neutral-100);
    border-radius: var(--radius-lg);
    padding: 20px;
    border: 1px solid var(--border-light);
}

.subject-group-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 12px 16px;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    margin-bottom: 16px;
    box-shadow: var(--shadow-sm);
}

.subject-group-header h5 {
    margin: 0;
    font-size: 1rem;
    font-weight: 600;
    color: var(--text-primary);
}

.subject-group-header .subject-count {
    font-size: 0.8rem;
    color: var(--text-secondary);
    background: var(--neutral-200);
    padding: 4px 10px;
    border-radius: 10px;
    font-weight: 500;
}

.subject-group-cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
    gap: 15px;
}

/* 컴팩트 테이블 스타일 */
.subject-group-section.compact {
    padding: 14px;
    margin-bottom: 0;
    height: fit-content;
}

.subject-group-section.compact .subject-group-header {
    margin-bottom: 12px;
    padding: 8px 12px;
}

.subject-group-section.compact .subject-group-header h5 {
    font-size: 0.95rem;
}

.subject-table {
    width: 100%;
    border-collapse: collapse;
    background: var(--bg-card);
    border-radius: var(--radius-md);
    overflow: hidden;
    font-size: 0.8rem;
}

.subject-table thead {
    background: linear-gradient(135deg, var(--neutral-200) 0%, var(--neutral-100) 100%);
}

.subject-table th {
    padding: 8px 6px;
    text-align: center;
    font-weight: 600;
    color: var(--text-primary);
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.3px;
    border-bottom: 2px solid var(--neutral-300);
}

.subject-table th:first-child {
    text-align: left;
    padding-left: 10px;
}

.subject-table td {
    padding: 8px 6px;
    border-bottom: 1px solid var(--neutral-200);
    color: var(--text-primary);
}

.subject-table td.center {
    text-align: center;
}

.subject-table td.subject-name-cell {
    font-weight: 500;
    padding-left: 10px;
    max-width: 100px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

.subject-table tbody tr:hover {
    background: var(--neutral-100);
}

.subject-table tbody tr:last-child td {
    border-bottom: none;
}

.subject-table tr.no-grade-row {
    opacity: 0.7;
    background: var(--neutral-50);
}

.subject-table .score-value {
    font-weight: 600;
    color: var(--text-primary);
}

.subject-table .avg-value {
    font-size: 0.75rem;
    color: var(--text-muted);
    margin-left: 2px;
}

.subject-table .achievement-badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 600;
    font-size: 0.8rem;
}

.subject-table .achievement-badge.A { background: var(--success); color: var(--text-inverse); }
.subject-table .achievement-badge.B { background: var(--info); color: var(--text-inverse); }
.subject-table .achievement-badge.C { background: var(--accent); color: var(--text-inverse); }
.subject-table .achievement-badge.D { background: var(--warning); color: var(--text-inverse); }
.subject-table .achievement-badge.E { background: var(--primary); color: var(--text-inverse); }

.subject-table .grade9-value {
    color: var(--accent);
    font-weight: 600;
}

@media (max-width: 768px) {
    .subject-table {
        font-size: 0.75rem;
    }

    .subject-table th,
    .subject-table td {
        padding: 8px 4px;
    }

    .subject-table td.subject-name-cell {
        max-width: 80px;
    }

    .subject-table .avg-value {
        display: none;
    }
}

.subject-card.no-grade {
    opacity: 0.8;
    border-left: 4px solid var(--neutral-500);
}

.subject-metrics.simple {
    grid-template-columns: 1fr 1fr;
    margin-bottom: 0;
}

.no-grade-notice {
    text-align: center;
    padding: 15px;
    background: var(--neutral-200);
    border-radius: var(--radius-sm);
    margin-top: 15px;
}

.no-grade-notice span {
    color: var(--text-secondary);
    font-size: 0.9rem;
    font-style: italic;
}

.subject-card {
    background: var(--bg-card);
    border-radius: var(--radius-md);
    padding: 18px;
    box-shadow: var(--shadow-sm);
    transition: all 0.2s ease;
    border: 1px solid var(--border-light);
}

.subject-card:hover {
    box-shadow: var(--shadow-md);
    border-color: var(--border-medium);
}

.subject-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 15px;
    padding-bottom: 10px;
    border-bottom: 2px solid var(--neutral-200);
}

.subject-header h5 {
    color: var(--text-primary);
    font-size: 1.1rem;
    font-weight: 600;
    margin: 0;
}

.subject-header .credits {
    background: var(--info);
    color: var(--text-inverse);
    padding: 4px 8px;
    border-radius: 12px;
    font-size: 0.8rem;
    font-weight: 500;
}

.subject-metrics {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 15px;
    margin-bottom: 15px;
}

.subject-metrics:last-of-type {
    grid-template-columns: 1fr 1fr 0fr;
}

.metric {
    text-align: center;
}

.metric-label {
    display: block;
    font-size: 0.8rem;
    color: var(--text-secondary);
    margin-bottom: 5px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.metric-value {
    display: block;
    font-size: 1rem;
    font-weight: 600;
    color: var(--text-primary);
}

.metric-average {
    display: block;
    font-size: 0.8rem;
    color: var(--text-secondary);
    font-weight: normal;
    margin-top: 2px;
}

.metric-value.achievement.A {
    background: var(--success);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.B {
    background: var(--info);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.C {
    background: var(--accent);
    color: var(--text-primary);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.D {
    background: var(--warning);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}
.metric-value.achievement.E, .metric-value.achievement.미도달 {
    background: var(--primary);
    color: var(--text-inverse);
    font-weight: bold;
    padding: 4px 8px;
    border-radius: 4px;
}

.percentile-bar {
    height: 8px;
    background: var(--neutral-200);
    border-radius: 4px;
    overflow: hidden;
    position: relative;
}

.percentile-fill {
    height: 100%;
    border-radius: 4px;
    transition: width 0.8s ease;
}

.percentile-fill.excellent { background: linear-gradient(90deg, var(--success), var(--success-light)); }
.percentile-fill.good { background: linear-gradient(90deg, var(--info), var(--info-light)); }
.percentile-fill.average { background: linear-gradient(90deg, var(--warning), var(--warning-light)); }
.percentile-fill.low { background: linear-gradient(90deg, var(--neutral-500), var(--neutral-400)); }

.percentile.excellent { color: var(--success); font-weight: 600; }
.percentile.good { color: var(--info); font-weight: 600; }
.percentile.average { color: var(--warning); font-weight: 600; }
.percentile.low { color: var(--neutral-500); font-weight: 500; }

@media (max-width: 1024px) {
    .analysis-overview {
        grid-template-columns: 1fr;
        gap: 20px;
    }
    
    .chart-container {
        padding: 20px;
    }
    
    .student-detail-header {
        flex-direction: column;
        gap: 20px;
    }
    
    .overall-stats {
        align-self: stretch;
        justify-content: space-around;
    }
    
    .subject-cards {
        grid-template-columns: 1fr;
    }
}

/* 출력용 스타일 */
@page {
    size: A4 portrait;
    margin: 10mm;
}
@media print {
    .print-area {
        transform-origin: top left !important;
    }
    .print-area.apply-print-scale {
        transform: scale(var(--page-scale, 1)) !important;
    }
    * {
        -webkit-print-color-adjust: exact !important;
        color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    body {
        background: white !important;
        margin: 0;
        padding: 0;
        font-size: 12px;
        line-height: 1.4;
        color: #000 !important;
    }
    
    .container {
        max-width: none;
        margin: 0;
        box-shadow: none;
        border-radius: 0;
        background: white;
    }
    
    header {
        display: none !important;
    }
    
    .upload-section,
    .tabs,
    .view-toggle,
    .student-selector,
    .search-box,
    .print-controls {
        display: none !important;
    }
    
    .results-section {
        padding: 15px;
    }
    
    .tab-content {
        display: block !important;
    }
    
    .tab-content:not(.print-target) {
        display: none !important;
    }

    /* 학생 탭 인쇄 시 개인 상세 페이지만 표시 */
    #students-tab.only-class-print > *:not(.class-print-area) {
        display: none !important;
    }

    /* A4에 맞춘 폭 고정 및 중앙 정렬 */
    .class-print-area {
        width: 190mm;
        margin: 0 auto;
    }
    .class-print-area .student-print-page {
        width: 190mm;
        transform-origin: top left !important;
    }
    .class-print-area .student-print-page.apply-print-scale {
        transform: scale(var(--page-scale, 1)) !important;
    }

    /* 학급 전체 인쇄 모드: 더 컴팩트한 카드와 차트 크기 */
    #students-tab.only-class-print .student-detail-header {
        padding: 12px;
        margin-bottom: 10px;
    }
    #students-tab.only-class-print .student-info h3 {
        font-size: 14px;
    }
    #students-tab.only-class-print .student-meta {
        font-size: 11px;
    }
    #students-tab.only-class-print .summary-card,
    #students-tab.only-class-print .stat-card {
        margin-bottom: 8px;
        padding: 10px;
    }
    
    /* 별도 프린트 헤더는 사용하지 않음 */
    .print-header {
        display: none !important;
    }
    
    .print-header h2 {
        margin: 0;
        color: #2c3e50;
        font-size: 18px;
        font-weight: bold;
    }
    
    .print-date {
        margin-top: 10px;
        font-size: 12px;
        color: #666;
    }
    
    .student-detail-header {
        background: #f8f9fa !important;
        border: 2px solid #4facfe;
        margin-bottom: 15px;
        padding: 20px;
        page-break-after: avoid;
    }
    
    .student-info h3 {
        color: #2c3e50 !important;
        font-size: 16px;
        margin-bottom: 8px;
    }
    
    .student-meta {
        color: #666 !important;
        font-size: 12px;
    }
    
    .stat-card {
        border: 1px solid #ddd !important;
        background: white !important;
    }
    
    .stat-label {
        color: #666 !important;
        font-size: 10px;
    }
    
    .stat-value {
        color: #2c3e50 !important;
        font-size: 14px;
    }
    
    .analysis-overview {
        grid-template-columns: 1fr;
        gap: 15px;
        page-break-inside: avoid;
    }
    
    /* 레이더 차트도 출력/PDF에 포함 */
    .chart-container {
        display: block !important;
    }
    
    .summary-card {
        border: 1px solid #ddd !important;
        background: #f9f9f9 !important;
        margin-bottom: 15px;
    }
    
    .summary-header h4 {
        color: #2c3e50 !important;
        font-size: 14px;
    }
    
    .summary-label {
        color: #666 !important;
        font-size: 11px;
    }
    
    .summary-value {
        color: #2c3e50 !important;
        font-size: 12px;
    }
    
    .summary-value.highlight {
        color: #4facfe !important;
        font-weight: bold;
    }
    
    .subject-details h4 {
        color: #2c3e50 !important;
        font-size: 14px;
        margin-bottom: 15px;
    }
    
    .subject-cards {
        grid-template-columns: repeat(2, 1fr);
        gap: 10px;
        page-break-inside: avoid;
    }
    
    .subject-card {
        border: 1px solid #ddd !important;
        background: white !important;
        page-break-inside: avoid;
        margin-bottom: 8px;
        padding: 12px;
    }
    
    .subject-header h5 {
        color: #2c3e50 !important;
        font-size: 12px;
        margin: 0 0 8px 0;
    }
    
    .subject-header .credits {
        background: #4facfe !important;
        color: white !important;
        font-size: 9px;
        padding: 2px 6px;
    }
    
    .subject-metrics {
        gap: 8px;
        margin-bottom: 8px;
    }
    
    .metric-label {
        font-size: 9px;
        color: #666 !important;
    }
    
    .metric-value {
        font-size: 11px;
        color: #2c3e50 !important;
    }
    
    .metric-average {
        font-size: 9px;
        color: #666 !important;
    }
    
    .percentile-bar {
        height: 6px;
        background: #e9ecef !important;
    }
    
    .no-grade-notice {
        background: rgba(108, 117, 125, 0.1) !important;
        font-size: 10px;
    }
    
    .no-grade-notice span {
        color: #666 !important;
    }
    
    /* 성취도 색상 */
    .achievement.A, .metric-value.achievement.A { 
        background: #28a745 !important; 
        color: white !important; 
    }
    .achievement.B, .metric-value.achievement.B { 
        background: #17a2b8 !important; 
        color: white !important; 
    }
    .achievement.C, .metric-value.achievement.C { 
        background: #ffc107 !important; 
        color: #212529 !important; 
    }
    .achievement.D, .metric-value.achievement.D { 
        background: #fd7e14 !important; 
        color: white !important; 
    }
    .achievement.E, .achievement.미도달, 
    .metric-value.achievement.E, .metric-value.achievement.미도달 { 
        background: #dc3545 !important; 
        color: white !important; 
    }
    
    /* 페이지 나누기 규칙 */
    .subject-card {
        break-inside: avoid;
    }
    
    .summary-card {
        break-inside: avoid;
    }

    /* 학급 전체 인쇄: 학생별 한 페이지씩 */
    .class-print-area .student-print-page {
        page-break-after: always;
        break-after: page;
    }
    .class-print-area .student-print-page:last-child {
        page-break-after: auto;
        break-after: auto;
    }
}

/* PDF 출력 버튼 스타일 */
.print-controls {
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
}

.student-nav-controls {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
}

.student-nav-status {
    color: var(--text-secondary);
    font-size: 0.9rem;
    font-weight: 600;
    padding: 0 4px;
}

.print-btn, .pdf-btn {
    background: linear-gradient(135deg, var(--success) 0%, var(--success-light) 100%);
    color: var(--text-inverse);
    border: none;
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    font-size: 0.9rem;
    cursor: pointer;
    transition: all 0.25s ease;
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 500;
}

.pdf-btn {
    background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
}

.print-btn:hover, .pdf-btn:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
}

.print-btn::before {
    content: "🖨️";
    font-size: 16px;
}

.pdf-btn::before {
    content: "📄";
    font-size: 16px;
}

@media (max-width: 768px) {
    .student-selector {
        flex-direction: column;
        align-items: stretch;
        gap: 15px;
    }

    .selector-group {
        justify-content: space-between;
    }

    .selector {
        min-width: unset;
        flex: 1;
    }

    .subject-metrics {
        grid-template-columns: repeat(2, 1fr);
        gap: 10px;
    }

    .container {
        margin: 10px;
        border-radius: var(--radius-lg);
    }

    header {
        padding: 20px;
    }

    header h1 {
        font-size: 1.6rem;
    }

    .header-subtitle {
        font-size: 0.92rem;
    }

    .upload-section,
    .results-section {
        padding: 20px;
    }

    .file-input-label {
        min-height: 74px;
        padding: 16px 18px;
    }

    .action-buttons {
        flex-direction: column;
        align-items: stretch;
    }

    .subject-averages {
        grid-template-columns: 1fr;
    }

    .tabs {
        display: flex;
        width: 100%;
    }

    .tab-btn {
        min-width: 0;
        padding: 12px 10px;
        font-size: 0.85rem;
    }

    .students-grid {
        grid-template-columns: 1fr;
        gap: 15px;
    }

    .student-card-header {
        padding: 14px 16px;
        gap: 8px;
    }

    .student-card-badges {
        gap: 6px;
    }

    .card-badge {
        padding: 4px 8px;
        gap: 4px;
    }

    .card-badge-label {
        font-size: 0.6rem;
    }

    .card-badge-value {
        font-size: 0.8rem;
    }

    .student-summary {
        gap: 5px;
    }

    .summary-metric-inline {
        padding: 3px 6px;
    }

    .summary-metric-inline .metric-label {
        font-size: 0.6rem;
    }

    .summary-metric-inline .metric-value {
        font-size: 0.78rem;
    }

    .subject-data {
        gap: 6px;
    }

    .subject-name {
        font-size: 0.85rem;
    }

    .subject-score, .subject-achievement, .subject-grade, .subject-percentile {
        font-size: 0.75rem;
    }

    .print-controls {
        justify-content: center;
    }

    .print-btn, .pdf-btn {
        flex: 1;
        min-width: 120px;
        justify-content: center;
    }

    .analyze-btn,
    .secondary-btn {
        width: 100%;
        justify-content: center;
    }

    .print-controls,
    .student-nav-controls {
        justify-content: center;
    }
}
`;
    }

    // JavaScript 파일 내용 가져오기 (실제 동작하는 버전)
    async getScriptJS() {
        return `
// 독립형 HTML용 ScoreAnalyzer 클래스
class StandaloneScoreAnalyzer {
    constructor() {
        this.combinedData = window.PRELOADED_DATA || null;
        this.initializeEventListeners();
        if (this.combinedData) {
            this.displayResults();
        }
    }

    initializeEventListeners() {
        // 탭 전환 기능
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                this.switchTab(e.target.getAttribute('data-tab'));
            });
        });

        // 학생 선택 기능들
        const studentSearch = document.getElementById('studentSearch');
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');

        if (gradeSelect) {
            gradeSelect.addEventListener('change', () => {
                this.updateClassOptions();
                this.updateStudentOptions();
                this.filterStudentTable();
            });
        }

        if (classSelect) {
            classSelect.addEventListener('change', () => {
                this.updateStudentOptions();
                this.filterStudentTable();
            });
        }

        if (studentNameSearch) {
            studentNameSearch.addEventListener('input', () => {
                this.updateStudentOptions();
            });
        }

        if (studentSearch) {
            studentSearch.addEventListener('input', () => {
                this.filterStudentTable();
            });
        }

        if (studentSelect) {
            studentSelect.addEventListener('change', () => {
                const showBtn = document.getElementById('showStudentDetail');
                if (showBtn) {
                    showBtn.disabled = !studentSelect.value;
                }
            });
        }

        // 상세 분석 버튼
        const showStudentDetail = document.getElementById('showStudentDetail');
        if (showStudentDetail) {
            showStudentDetail.addEventListener('click', () => {
                this.showStudentDetail();
            });
        }

        // 뷰 전환 버튼들
        const tableViewBtn = document.getElementById('tableViewBtn');
        const detailViewBtn = document.getElementById('detailViewBtn');

        if (tableViewBtn) {
            tableViewBtn.addEventListener('click', () => {
                this.switchView('table');
            });
        }

        if (detailViewBtn) {
            detailViewBtn.addEventListener('click', () => {
                this.switchView('detail');
            });
        }
    }

    switchTab(tabName) {
        // 모든 탭 버튼과 콘텐츠 비활성화
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.style.display = 'none');
        
        // 선택된 탭 활성화
        const tabBtn = document.querySelector('[data-tab="' + tabName + '"]');
        const tabContent = document.getElementById(tabName + '-tab');
        
        if (tabBtn) tabBtn.classList.add('active');
        if (tabContent) tabContent.style.display = 'block';
    }

    displayResults() {
        if (!this.combinedData) return;
        
        this.displaySubjectAverages();
        this.displayGradeAnalysis();
        this.displayStudentAnalysis();
        if (document.querySelector('[data-tab="subjects"]') && document.getElementById('subjects-tab')) {
            this.switchTab('subjects');
        }
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
                    distributionHTML += \`
                        <div class="achievement-bar">
                            <span class="achievement-label">\${grade}</span>
                            <div class="achievement-bar-container">
                                <div class="achievement-bar-fill" style="width: \${percentage}%"></div>
                            </div>
                            <span class="achievement-percentage">\${percentage.toFixed(1)}%</span>
                        </div>
                    \`;
                });
                distributionHTML += '</div>';
            }
            
            subjectDiv.innerHTML = \`
                <div class="subject-header">
                    <h3>\${subject.name}</h3>
                    <span class="credits">\${subject.credits || 0}학점</span>
                </div>
                <div class="average-score">
                    <span class="score">\${subject.average?.toFixed(1) || 'N/A'}</span>
                    <span class="label">평균 점수</span>
                </div>
                \${distributionHTML}
            \`;
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

    displayStudentAnalysis() {
        if (!this.combinedData) return;

        this.populateStudentSelectors();
        const container = document.getElementById('studentTable');
        this.renderStudentTable(this.combinedData.students, this.combinedData.subjects, container);
    }

    populateStudentSelectors() {
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        
        if (!gradeSelect || !classSelect) return;
        
        // 학년 옵션 생성
        const grades = [...new Set(this.combinedData.students.map(s => s.grade).filter(g => g))].sort();
        gradeSelect.innerHTML = '<option value="">전체</option>';
        grades.forEach(grade => {
            const option = document.createElement('option');
            option.value = grade;
            option.textContent = grade + '학년';
            gradeSelect.appendChild(option);
        });

        // 반 옵션 생성
        const classes = [...new Set(this.combinedData.students.map(s => s.class).filter(c => c))].sort();
        classSelect.innerHTML = '<option value="">전체</option>';
        classes.forEach(cls => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = cls + '반';
            classSelect.appendChild(option);
        });

        this.updateStudentOptions();
    }

    updateClassOptions() {
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        if (!gradeSelect || !classSelect) return;
        
        const selectedGrade = gradeSelect.value;

        let students = this.combinedData.students;
        if (selectedGrade) {
            students = students.filter(s => s.grade == selectedGrade);
        }

        const classes = [...new Set(students.map(s => s.class).filter(c => c))].sort();
        classSelect.innerHTML = '<option value="">전체</option>';
        classes.forEach(cls => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = cls + '반';
            classSelect.appendChild(option);
        });
    }

    updateStudentOptions() {
        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSelect = document.getElementById('studentSelect');
        const studentNameSearch = document.getElementById('studentNameSearch');
        
        if (!studentSelect) return;
        
        const selectedGrade = gradeSelect ? gradeSelect.value : '';
        const selectedClass = classSelect ? classSelect.value : '';
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
            students = students.filter(s => 
                (s.name && s.name.toLowerCase().includes(q)) || 
                (s.originalNumber && String(s.originalNumber).includes(q))
            );
        }

        studentSelect.innerHTML = '<option value="">학생 선택</option>';
        students.forEach(student => {
            const option = document.createElement('option');
            option.value = student.number || student.originalNumber;
            option.textContent = (student.originalNumber || student.number || '') + '번 - ' + (student.name || '');
            studentSelect.appendChild(option);
        });

        const showBtn = document.getElementById('showStudentDetail');
        if (showBtn) {
            showBtn.disabled = students.length !== 1 && !studentSelect.value;
        }
    }

    renderStudentTable(students, subjects, container) {
        if (!container) return;
        
        container.innerHTML = '';

        if (students.length === 0) {
            container.innerHTML = '<p>학생 데이터가 없습니다.</p>';
            return;
        }

        // 테이블 헤더 생성
        const headerRow = ['번호', '이름', '평균등급'];
        subjects.forEach(subject => {
            headerRow.push(subject.name);
        });

        let tableHTML = '<table><thead><tr>';
        headerRow.forEach(header => {
            tableHTML += '<th>' + header + '</th>';
        });
        tableHTML += '</tr></thead><tbody>';

        // 학생 데이터 행 생성
        students.forEach(student => {
            tableHTML += '<tr>';
            tableHTML += '<td>' + (student.originalNumber || student.number || '') + '</td>';
            tableHTML += '<td>' + (student.name || '') + '</td>';
            tableHTML += '<td>' + (student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : '-') + '</td>';
            
            subjects.forEach(subject => {
                const grade = student.grades ? student.grades[subject.name] : '';
                tableHTML += '<td>' + (grade || '-') + '</td>';
            });
            
            tableHTML += '</tr>';
        });

        tableHTML += '</tbody></table>';
        container.innerHTML = tableHTML;
    }

    filterStudentTable() {
        if (!this.combinedData) return;

        const gradeSelect = document.getElementById('gradeSelect');
        const classSelect = document.getElementById('classSelect');
        const studentSearch = document.getElementById('studentSearch');

        const selectedGrade = gradeSelect ? gradeSelect.value : '';
        const selectedClass = classSelect ? classSelect.value : '';
        const searchTerm = studentSearch ? studentSearch.value.trim().toLowerCase() : '';

        // 학년/반/검색어로 필터링
        let filtered = this.combinedData.students;

        if (selectedGrade) {
            filtered = filtered.filter(s => String(s.grade) === String(selectedGrade));
        }

        if (selectedClass) {
            filtered = filtered.filter(s => String(s.class) === String(selectedClass));
        }

        if (searchTerm) {
            filtered = filtered.filter(s =>
                s.number.toString().includes(searchTerm) ||
                s.name.toLowerCase().includes(searchTerm)
            );
        }

        // 테이블 다시 렌더링
        const container = document.getElementById('studentTable');
        if (container) {
            this.renderStudentTable(filtered, this.combinedData.subjects, container);
        }
    }

    // 뷰 전환 기능
    switchView(viewType) {
        const tableViewBtn = document.getElementById('tableViewBtn');
        const detailViewBtn = document.getElementById('detailViewBtn');
        const tableView = document.getElementById('tableView');
        const detailView = document.getElementById('detailView');

        if (viewType === 'table') {
            if (tableViewBtn) tableViewBtn.classList.add('active');
            if (detailViewBtn) detailViewBtn.classList.remove('active');
            if (tableView) tableView.style.display = 'block';
            if (detailView) detailView.style.display = 'none';
        } else {
            if (tableViewBtn) tableViewBtn.classList.remove('active');
            if (detailViewBtn) detailViewBtn.classList.add('active');
            if (tableView) tableView.style.display = 'none';
            if (detailView) detailView.style.display = 'block';
        }
    }

    // 학생 상세 보기
    showStudentDetail() {
        const studentSelect = document.getElementById('studentSelect');
        const selectedStudentId = studentSelect ? studentSelect.value : '';
        
        if (!selectedStudentId) return;

        const student = this.combinedData.students.find(s => 
            (s.number && s.number == selectedStudentId) || 
            (s.originalNumber && s.originalNumber == selectedStudentId)
        );
        
        if (!student) return;

        this.renderStudentDetail(student);
        this.switchView('detail');
    }

    // 학생 상세 정보 렌더링
    renderStudentDetail(student) {
        const container = document.getElementById('studentDetailContent');
        if (!container) return;
        
        // 평균등급 순위 계산
        const studentsWithGrades = this.combinedData.students.filter(s => s.weightedAverageGrade);
        studentsWithGrades.sort((a, b) => a.weightedAverageGrade - b.weightedAverageGrade);
        
        const studentRank = studentsWithGrades.findIndex(s => s.number === student.number || s.originalNumber === student.originalNumber) + 1;
        const totalGradedStudents = studentsWithGrades.length;
        
        // 같은 등급 학생 수 계산
        const sameGradeStudents = studentsWithGrades.filter(s => 
            Math.abs(s.weightedAverageGrade - student.weightedAverageGrade) < 0.01
        );
        const sameGradeCount = sameGradeStudents.length;

        const html = \`
            <div class="student-detail-header">
                <div class="student-info">
                    <h3>\${student.name || '이름 없음'}</h3>
                    <div class="student-meta">
                        <span class="grade-class">\${student.grade || ''}학년 \${student.class || ''}반 \${student.originalNumber || student.number || ''}번</span>
                        \${student.fileName ? \`<span class="file-info">출처: \${student.fileName}</span>\` : ''}
                    </div>
                </div>
                <div class="overall-stats">
                    <div class="stat-card">
                        <span class="stat-label">평균등급</span>
                        <span class="stat-value grade">\${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                    </div>
                    <div class="stat-card">
                        <span class="stat-label">전체 학생수</span>
                        <span class="stat-value">\${totalGradedStudents}명</span>
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
                                    <span class="summary-value">\${student.grade || ''}학년 \${student.class || ''}반 \${student.originalNumber || student.number || ''}번</span>
                                </div>
                                <div class="summary-item">
                                    <span class="summary-label">평균등급</span>
                                    <span class="summary-value highlight">\${student.weightedAverageGrade ? student.weightedAverageGrade.toFixed(2) : 'N/A'}</span>
                                </div>
                                <div class="summary-item">
                                    <span class="summary-label">평균등급(9등급환산)</span>
                                    <span class="summary-value orange">\${student.weightedAverage9Grade ? student.weightedAverage9Grade.toFixed(2) : 'N/A'}</span>
                                </div>
                                <div class="summary-item">
                                    <span class="summary-label">등급 순위</span>
                                    <span class="summary-value highlight">\${studentRank}/\${totalGradedStudents}위\${sameGradeCount > 1 ? \` (\${sameGradeCount}명)\` : ''}</span>
                                </div>
                                <div class="summary-item">
                                    <span class="summary-label">전체 학생수</span>
                                    <span class="summary-value">\${totalGradedStudents}명</span>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="chart-container">
                        <h4>교과(군)별 평균등급</h4>
                        <canvas id="studentPercentileChart" width="400" height="400"></canvas>
                    </div>
                </div>

                <div class="subject-details">
                    <h4>과목별 상세 분석</h4>
                    <div class="subject-cards">
                        \${this.renderSubjectCards(student)}
                    </div>
                </div>
            </div>
        \`;
        
        container.innerHTML = html;
        
        // 학생 차트 생성
        setTimeout(() => {
            this.createStudentPercentileChart(student);
        }, 100);
    }

    // 과목별 카드 렌더링
    renderSubjectCards(student) {
        if (!student.grades || !this.combinedData.subjects) return '';
        
        return this.combinedData.subjects.map(subject => {
            const grade = student.grades[subject.name];
            if (!grade) return '';
            
            // 해당 과목에서의 순위 계산
            const subjectStudents = this.combinedData.students
                .filter(s => s.grades && s.grades[subject.name])
                .sort((a, b) => a.grades[subject.name] - b.grades[subject.name]);
            
            const subjectRank = subjectStudents.findIndex(s => 
                (s.number === student.number || s.originalNumber === student.originalNumber)
            ) + 1;
            
            return \`
                <div class="subject-card detailed">
                    <div class="subject-header">
                        <h5>\${subject.name}</h5>
                        <div class="subject-grade grade-\${Math.ceil(grade)}">\${grade}등급</div>
                    </div>
                    <div class="subject-stats">
                        <div class="stat-item">
                            <span class="stat-label">등급</span>
                            <span class="stat-value">\${grade}등급</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">과목내 순위</span>
                            <span class="stat-value">\${subjectRank}/\${subjectStudents.length}위</span>
                        </div>
                    </div>
                </div>
            \`;
        }).filter(card => card).join('');
    }

    // 산점도 차트 생성
    createScatterChart(students) {
        const ctx = document.getElementById('scatterChart');
        if (!ctx) return;
        
        const canvas = ctx.getContext ? ctx.getContext('2d') : null;
        if (!canvas) return;
        
        // 기존 차트가 있다면 파괴 및 동일 캔버스 잔존 차트 제거
        try { if (this.scatterChart) this.scatterChart.destroy(); } catch(_) {}
        try {
            const existing = (Chart.getChart ? Chart.getChart(canvas.canvas) : (canvas.canvas && (canvas.canvas._chart || canvas.canvas.chart)));
            if (existing && typeof existing.destroy === 'function') existing.destroy();
        } catch (_) {}

        // 평균등급별로 학생을 정렬
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

        // 산점도 데이터 생성
        const scatterData = [];
        const colors = ['#e74c3c', '#f39c12', '#f1c40f', '#2ecc71', '#3498db'];
        
        Object.keys(gradeGroups).forEach(grade => {
            const studentsInGrade = gradeGroups[grade];
            studentsInGrade.forEach((student, index) => {
                const gradeNum = parseFloat(grade);
                const colorIndex = Math.min(Math.floor(gradeNum), 4);
                scatterData.push({
                    x: gradeNum,
                    y: index + 1,
                    backgroundColor: colors[colorIndex],
                    borderColor: colors[colorIndex],
                    studentName: student.name
                });
            });
        });

        this.scatterChart = new Chart(canvas, {
            type: 'scatter',
            data: {
                datasets: [{
                    label: '학생 분포',
                    data: scatterData,
                    backgroundColor: scatterData.map(d => d.backgroundColor),
                    borderColor: scatterData.map(d => d.borderColor),
                    pointRadius: 6,
                    pointHoverRadius: 8
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const point = context.raw;
                                return point.studentName + ': ' + point.x + '등급';
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: '평균등급'
                        },
                        min: 1,
                        max: 5,
                        reverse: false
                    },
                    y: {
                        title: {
                            display: true,
                            text: '학생 수'
                        },
                        min: 0,
                        beginAtZero: true
                    }
                }
            }
        });
    }

    // 등급 분포 막대차트 생성
    createGradeDistributionChart(students) {
        const ctx = document.getElementById('barChart');
        if (!ctx) return;
        
        const canvas = ctx.getContext ? ctx.getContext('2d') : null;
        if (!canvas) return;
        
        // 기존 차트가 있다면 파괴 및 동일 캔버스 잔존 차트 제거
        try { if (this.barChart) this.barChart.destroy(); } catch(_) {}
        try {
            const existing = (Chart.getChart ? Chart.getChart(canvas.canvas) : (canvas.canvas && (canvas.canvas._chart || canvas.canvas.chart)));
            if (existing && typeof existing.destroy === 'function') existing.destroy();
        } catch (_) {}

        // 등급별 구간 정의
        const gradeRanges = [
            { label: '1.0-1.5', min: 1.0, max: 1.5, color: '#e74c3c' },
            { label: '1.5-2.0', min: 1.5, max: 2.0, color: '#e67e22' },
            { label: '2.0-2.5', min: 2.0, max: 2.5, color: '#f39c12' },
            { label: '2.5-3.0', min: 2.5, max: 3.0, color: '#f1c40f' },
            { label: '3.0-3.5', min: 3.0, max: 3.5, color: '#2ecc71' },
            { label: '3.5-4.0', min: 3.5, max: 4.0, color: '#27ae60' },
            { label: '4.0-4.5', min: 4.0, max: 4.5, color: '#3498db' },
            { label: '4.5-5.0', min: 4.5, max: 5.0, color: '#2980b9' }
        ];

        const rangeCounts = gradeRanges.map(range => {
            return students.filter(student => 
                student.weightedAverageGrade >= range.min && 
                student.weightedAverageGrade < range.max
            ).length;
        });

        this.barChart = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: gradeRanges.map(range => range.label),
                datasets: [{
                    label: '학생 수',
                    data: rangeCounts,
                    backgroundColor: gradeRanges.map(range => range.color),
                    borderColor: gradeRanges.map(range => range.color),
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '학생 수'
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: '평균등급 구간'
                        }
                    }
                }
            }
        });
    }

    // 학생 레이더 차트 생성
    createStudentPercentileChart(student) {
        const ctx = document.getElementById('studentPercentileChart');
        if (!ctx) return;
        
        const canvas = ctx.getContext ? ctx.getContext('2d') : null;
        if (!canvas) return;
        
        // 기존 차트 제거 및 동일 캔버스의 잔존 차트 제거
        try { if (this.studentPercentileChart) this.studentPercentileChart.destroy(); } catch(_) {}
        try {
            const existing = (Chart.getChart ? Chart.getChart(canvas.canvas) : (canvas.canvas && (canvas.canvas._chart || canvas.canvas.chart)));
            if (existing && typeof existing.destroy === 'function') existing.destroy();
        } catch (_) {}

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
            return grade ? (6 - grade) : 0; // 등급을 역산하여 높을수록 좋게
        });

        this.studentPercentileChart = new Chart(canvas, {
            type: 'radar',
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
                    pointHoverBackgroundColor: '#fff',
                    pointHoverBorderColor: 'rgba(52, 152, 219, 1)'
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    r: {
                        angleLines: {
                            display: true
                        },
                        grid: {
                            circular: true
                        },
                        pointLabels: {
                            display: true,
                            centerPointLabels: true,
                            font: {
                                size: 12
                            }
                        },
                        ticks: {
                            display: true,
                            stepSize: 1,
                            min: 0,
                            max: 5,
                            callback: function(value) {
                                return (6 - value) + '등급';
                            }
                        }
                    }
                }
            }
        });
    }
}

// 페이지 로드 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    new StandaloneScoreAnalyzer();
});
        `;
    }
}

// 전역 변수로 선언
let scoreAnalyzer;

// 페이지 로드 시 분석기 초기화
document.addEventListener('DOMContentLoaded', () => {
    scoreAnalyzer = new ScoreAnalyzer();
});
