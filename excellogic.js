const table = document.querySelector('#excel_table');
const Excel = require('exceljs');
const { ipcRenderer } = require('electron');
var workbook = new Excel.Workbook();
    const form = document.querySelector("#excel_file");
    form.addEventListener('change',submitForm);
    function submitForm(e){
        document.getElementById('email_column_sel').disabled = false;
        table.innerHTML = '';
        e.preventDefault();
        const item = document.querySelector('#excel_file').files[0].path;
        console.log(item);
        document.querySelector('#excel-div').classList.remove("d-none")
        document.querySelector('#buttons').classList.remove("d-none")
        workbook.xlsx.readFile(item)
        .then(function() {
            var ws = workbook.getWorksheet(1);
            var cell = ws.getCell('A1').value;
            console.log(cell);
            rowCount = ws.actualRowCount;
            ws.eachRow(function(row, rowNumber) {
                // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
                const tr = document.createElement('tr');
                var RowArray = row.values//JSON.stringify(row.values);
                if(rowNumber==1){
                    var tr_con = "<td style='font-weight:bold; position:sticky;top:0;background-color:#f2f5fa;'>Sr. no</td>";
                    var sel1 = document.getElementById('email_column_sel');
                        sel1.innerHTML = '';
                    RowArray.forEach((item,index)=>{
                        tr_con += "<td style='font-weight:bold; position:sticky;top:0;background-color:#f2f5fa;'>"+item+"</td>";
                        rowValHeader.push({'text':item, 'value':index});
                        sel1.appendChild(new Option(item,index));
                    })
                    tr_con += "<td style='font-weight:bold; position:sticky;top:0;z-index:2;right:0;background-color:#f2f5fa;'>Select</td>";
                }else{
                    var tr_con = "<td>"+(rowNumber-1)+"</td>";
                    RowArray.forEach((item,index)=>{
                        tr_con += "<td>"+item+"</td>";
                    })
                    tr_con += `<td style='position:sticky;  text-align:center;z-index:1;right:0;background-color:rgba(200,200,200,0.5);'>
                        <input class='form-check-input' type='checkbox' id='sel_em[${(rowNumber-1)}]' name='sel_em[]' checked />
                    </td>`;
                }
                // console.log(tr_con);
                tr.innerHTML= tr_con;
                // tr.addEventListener('click',checkone);
                tr.addEventListener('click', myFunc, false);
                tr.index_no = rowNumber-2;
                tr.setAttribute("name", "rows[]"); 
                table.appendChild(tr);
            });
            document.addEventListener('DOMContentLoaded', function() {
                var elems = document.querySelectorAll('email_column_sel');
                var instances = M.FormSelect.init(elems,rowValHeader);
            });  
        });
    }

    document.getElementById('select-all').onclick = function() {
        var checkboxes = document.getElementsByName('sel_em[]');
        for (var checkbox of checkboxes) {
          checkbox.checked = true;
        }
    }
    document.getElementById('de-select-all').onclick = function() {
        var checkboxes = document.getElementsByName('sel_em[]');
        for (var checkbox of checkboxes) {
          checkbox.checked = false;
        }
    }
    // document.getElementById('subcheck').onclick = function() {
    //     var checkboxes = document.getElementsByName('sel_em[]');
    //     var res = [];
    //     for(let checkbox of checkboxes){
    //         res.push(checkbox.checked);
    //     }
    //     alert(res);
    // }
    document.getElementById('lockbutton').onclick = function() {
        console.log(document.getElementsByName('sel_em[]')[1].checked);
        var lock = document.getElementById('locklayer');
        var selectall = document.getElementById('select-all');
        var deselectall = document.getElementById('de-select-all');
        var excel_div = document.getElementById('excel-div');
        if(lock.classList.contains('d-none')){
            excel_div.classList.add('unscrollable');
            excel_div.classList.remove('scrollable');
            lock.classList.remove('d-none');
            selectall.classList.add('d-none');
            deselectall.classList.add('d-none');
            this.innerText = 'Unlock';
        }else{
            excel_div.classList.remove('unscrollable');
            excel_div.classList.add('scrollable');
            lock.classList.add('d-none');
            selectall.classList.remove('d-none');
            deselectall.classList.remove('d-none');
            this.innerText = 'Lock';
        }
    }
    var myFunc = function(evt) {
        var checkbox = document.getElementsByName('sel_em[]')[evt.currentTarget.index_no];
        checkbox.checked = checkbox.checked? false:true;
    }
    ipcRenderer.on('openexcel',function(){
        document.querySelector("#excel_file").click();
    })

    ipcRenderer.on('email_status',function(e,item){
        if(item.rowNumber == rowCount){
            loadingSpinner(false);
        }
        var row = document.getElementsByName('rows[]')[item.rowNumber-1];
        if(item.status){
            row.classList.add('greenbg');
        }else{
            console.log(item.status,item.rowNumber);
            console.log(item);
            row.classList.add('redbg');
        }
    })
    
    // loadingSpinner(true);
    //setTimeout(loadingSpinner,1000,false);