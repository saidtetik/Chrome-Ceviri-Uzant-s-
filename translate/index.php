
<?php

  if($_FILES){
  include "upload.php";
  }
  ?>

<!DOCTYPE html>
<html lang="en" dir="ltr">
  <head>
    <meta charset="utf-8">
    <title>Translate</title>
    <link rel="stylesheet" href="style.css">


    </head>
  <body>


<div class="baslik">
    АУДАРУ
</div>

<div id="drop_zone">
      <img src="128.png" alt="Drag file to here" title="Drag file to here!" draggable="false"><br>
      Drop your files here
</div>

<div class="form">
    <form action="" method="post" enctype="multipart/form-data">
        <p>  <input type="file" name="file" value="Choose a File" class="button">
          <input type="submit" name="" value="Upload" class="button"><br><br>
          <div><?php if(!empty($file_name)){  echo '<a  href="uploads/'.$file_name.'"> <input type="button" value = "Download" ></a>';}?></div>
        </p>
    </form>
</div>
<script  type="text/javascript" src="upload.js"></script>

  </body>
</html>
