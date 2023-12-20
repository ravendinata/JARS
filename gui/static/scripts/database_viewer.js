// On page load
$(document).ready(function() {
    var tbl = $('#datatable').DataTable(
        {
            autoWidth: true,
            lengthMenu: [[10, 25, 50, 100, -1], [10, 25, 50, 100, "All"]],
            pageLength: 10,
            buttons: [
                'copy', 
                {
                    extend: 'csv', 
                    title: '<!-- REPORT NAME -->',
                    messageTop: 'This report was generated by JARS. All data is confidential and should not be shared.'
                },
                {
                    extend: 'excel',
                    title: '<!-- REPORT NAME -->',
                    messageTop: 'This report was generated by JARS. All data is confidential and should not be shared.'
                },
                {
                    extend: 'pdf',
                    title: '<!-- REPORT NAME -->',
                    messageTop: 'This report was generated by JARS. All data is confidential and should not be shared.'
                },
                'print',
                'colvis'
            ],
            pagingType: "full_numbers",
            // scrollX: true,
            "deferRender": true
        }
    );

    tbl.buttons().container().appendTo('#datatable_wrapper .col-md-6:eq(0)', tbl.table().container());
    
    if (document.documentElement.getAttribute('data-bs-theme') == 'dark')
        document.getElementById('btnToggleColorMode').innerHTML = 'Switch to Light Mode';
    else
        document.getElementById('btnToggleColorMode').innerHTML = 'Switch to Dark Mode';
});

// Toggle color mode
document.getElementById('btnToggleColorMode').addEventListener('click',()=>{
    if (document.documentElement.getAttribute('data-bs-theme') == 'dark')
    {
        document.documentElement.setAttribute('data-bs-theme','light')
        document.getElementById('btnToggleColorMode').innerHTML = 'Switch to Dark Mode';
    }
    else
    {
        document.documentElement.setAttribute('data-bs-theme','dark')
        document.getElementById('btnToggleColorMode').innerHTML = 'Switch to Light Mode';
    }
});