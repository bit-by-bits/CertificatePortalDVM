let sheet = document.forms['getCerti']['excel_file'];

function validateForm(){
    if(sheet.value == null ){
        alert("Please upload a valid Excel file before generating certificates.");
    }
}
