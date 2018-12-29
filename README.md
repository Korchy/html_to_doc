# html_to_doc
HTML to DOC converter

Supports embeding images to the destination doc document.

Installation
-
- Copy ExportToWord.inc.php to some directory of your project

Using
-
    require_once(dirname(dirname(__FILE__)).'/ExportToWord.inc.php');

    $html = '<html><body><div class = "test">Test</div></body></html>';
    $css = '<style type = "text/css">.test {font-weight: 600;}</style>';
    $fileName = 'c:/test.doc';
    ExportToWord::htmlToDoc($html, $css, $fileName);

Source
-
This class based on info ftom the article https://habr.com/post/168977/
