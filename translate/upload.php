<?php


    $file_name = $_FILES['file']['name'];
    $file_tmp_name = $_FILES['file']['tmp_name'];
    $extensions =array("doc","docx");
    $extension =explode(".",$file_name);
    $type = strtolower(end($extension));
    
    if(in_array($type,$extensions) && move_uploaded_file($file_tmp_name, "uploads/".$file_name)){

        $word = new COM('Word.Application');
        $word->Documents->Open(realpath("uploads/".$file_name));
        $word->Run("Normal.Translate.translate");
        $word->ActiveDocument->Save();
        $word->ActiveDocument->Close();
        $word->Quit();



    }
    else{
      echo "<h1>Dosya Word Belgesi değil</h1>";
    }


 ?>
