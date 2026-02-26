//Student Data 
const courseData = [
    {
        studentId: "STU001",
        studentName: "Ali Khan",
        course: "Advanced Excel & VBA",
        instructor: "Dr. Ayesha Siddiqui",
        status: "Active",
        enrollmentDate: "2024-01-15",
        email: "ali.khan@email.com",
        phone: "+92-300-0101010"
    },
    {
        studentId: "STU002",
        studentName: "Sara Malik",
        course: "Office Automation Mastery",
        instructor: "Prof. Muhammad Hassan",
        status: "Active",
        enrollmentDate: "2024-01-20",
        email: "sara.malik@email.com",
        phone: "+92-300-0101020"
    },
    {
        studentId: "STU003",
        studentName: "Bilal Ahmed",
        course: "Data Analysis with Excel",
        instructor: "Dr. Ayesha Siddiqui",
        status: "Completed",
        enrollmentDate: "2023-11-10",
        email: "bilal.ahmed@email.com",
        phone: "+92-300-0101030"
    },
    {
        studentId: "STU004",
        studentName: "Hira Shah",
        course: "Advanced Excel & VBA",
        instructor: "Prof. Muhammad Hassan",
        status: "Active",
        enrollmentDate: "2024-02-01",
        email: "hira.shah@email.com",
        phone: "+92-300-0101040"
    },
    {
        studentId: "STU005",
        studentName: "Owais Raza",
        course: "Office Automation Mastery",
        instructor: "Dr. Ayesha Siddiqui",
        status: "Pending",
        enrollmentDate: "2024-02-15",
        email: "owais.raza@email.com",
        phone: "+92-300-0101050"
    },
    {
        studentId: "STU006",
        studentName: "Maryam Khan",
        course: "Data Analysis with Excel",
        instructor: "Prof. Muhammad Hassan",
        status: "Active",
        enrollmentDate: "2024-01-25",
        email: "maryam.khan@email.com",
        phone: "+92-300-0101060"
    },
];

// Calculate and display statistics
        function updateStats() {
            document.getElementById('totalStudents').textContent = courseData.length;
            
            const uniqueCourses = [...new Set(courseData.map(d => d.course))];
            document.getElementById('totalCourses').textContent = uniqueCourses.length;
            
            const avgGrade = Math.round(courseData.reduce((sum, d) => sum + d.grade, 0) / courseData.length);
            document.getElementById('avgGrade').textContent = avgGrade + '%';
            
            const activeCount = courseData.filter(d => d.status === 'Active').length;
            document.getElementById('activeCount').textContent = activeCount;
        }

        // Render table preview
        function renderTable() {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = courseData.map(student => `
                <tr>
                    <td><strong>${student.studentId}</strong></td>
                    <td>${student.studentName}</td>
                    <td>${student.course}</td>
                    <td>${student.instructor}</td>
                    <td><strong>${student.grade}%</strong></td>
                    <td><span class="badge badge-${student.status.toLowerCase()}">${student.status}</span></td>
                    <td>${student.enrollmentDate}</td>
                </tr>
            `).join('');
        }

        // Export to Excel using XLSX.js
        function exportToExcel() {
            const btn = document.querySelector('.export-btn');
            const btnText = document.getElementById('btnText');
            
            // Show loading state
            btnText.textContent = 'Generating...';
            btn.style.opacity = '0.8';
            
            setTimeout(() => {
                try {
                    // Create worksheet from data
                    const ws = XLSX.utils.json_to_sheet(courseData);
                    
                    // Set column widths
                    const wscols = [
                        {wch: 10}, {wch: 20}, {wch: 25}, {wch: 20}, 
                        {wch: 8}, {wch: 12}, {wch: 15}, {wch: 25}, {wch: 15}
                    ];
                    ws['!cols'] = wscols;
                    
                    // Create workbook
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Student Records");
                    
                    // Generate filename with current date
                    const date = new Date().toISOString().split('T')[0];
                    const filename = `Aptech_Student_Data_${date}.xlsx`;
                    
                    // Save file
                    XLSX.writeFile(wb, filename);
                    
                    // Show success state
                    btn.classList.add('success');
                    btnText.textContent = 'Downloaded!';
                    setTimeout(() => {
                        btn.classList.remove('success');
                        btnText.textContent = 'Export to Excel';
                        btn.style.opacity = '1';
                    }, 2000);
                    
                } catch (error) {
                    console.error('Export failed:', error);
                    btnText.textContent = 'Error!';
                    setTimeout(() => {
                        btnText.textContent = 'Export to Excel';
                        btn.style.opacity = '1';
                    }, 2000);
                }
            }, 500);
        }

        // Initialize
        document.addEventListener('DOMContentLoaded', () => {
            updateStats();
            renderTable();
        });