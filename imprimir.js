function printPage(id){
    document.getElementById("'" + id + "'").contentWindow.focus();
    document.getElementById("'" + id + "'").contentWindow.print(); 
}