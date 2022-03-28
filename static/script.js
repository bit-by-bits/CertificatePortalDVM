let sheet = document.forms['getCerti']['excel_file'];

function validateForm() {
    if (sheet.value === '') {
        alert("Please upload a valid Excel file before generating certificates.");
        return false;
    }
    else{
        return true;
    }
}
