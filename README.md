# html_to_doc
HTML to DOC converter.

This php class enables to convert HTML document to DOC format.

Supports embeding images to the destination doc document.

Installation
-
- Copy ExportToWord.inc.php to some directory of your project.
- Change $imgRef and $imgDir static variables in the ExportToWord.inc.php file with your project images paths.

Using
-
- Use require_once to include ExportToWord class to your project file.
- Call ExportToWord::htmlToDoc function from your project file to convert HTML to DOC.

Sample code:

    require_once(dirname(dirname(__FILE__)).'/ExportToWord.inc.php');

    $html = '<html><body><div class = "test">Test</div></body></html>';
    $css = '<style type = "text/css">.test {font-weight: 600;}</style>';
    $fileName = 'c:/test.doc';
    ExportToWord::htmlToDoc($html, $css, $fileName);

Sample files
-
test.php - file with sample code of using HTML - DOC converter

test.html - file with sample html body. The same html body used in the test.php file.

test.doc - converted to DOC html file (with test.html body)

Source
-
This class based on info from the article https://habr.com/post/168977/
