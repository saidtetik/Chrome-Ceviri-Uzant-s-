

var dropZone = document.getElementById('drop_zone');
//console.log("said");

dropZone.addEventListener("drop",function(e){
  e.preventDefault();
  e.stopPropagation();

  var xhr = new XMLHttpRequest();
  var formData = new FormData();

  xhr.open("POST","upload.php",true);
  formData.append('file',e.dataTransfer.files[0]);
  xhr.send(formData);

});
dropZone.addEventListener("dragover",function(e){
  e.preventDefault();
  e.stopPropagation();
  //console.log("said");

  this.style.borderColor = "black";
  this.style.backgroundColor = "#dadae2";

});
dropZone.addEventListener("dragleave",function(){
  this.style.borderColor = "#ccc";
  this.style.backgroundColor = "#ddd";
});
