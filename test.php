<?php
require_once(str_replace(array('/', '\\'), '/', realpath(__DIR__)).'/ExportToWord.inc.php');

$html = '<html><body><div class = "test">Test</div><img src = "/img/testimg.jpg"></body></html>';
$css = '<style type = "text/css">.test {font-weight: 600;}</style>';
$fileName = str_replace(array('/', '\\'), '/', realpath(__DIR__)).'/test.doc';
ExportToWord::htmlToDoc($html, $css, $fileName);
?>
