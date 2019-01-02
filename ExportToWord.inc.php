<?php
//------------------------------------------------------------------------------------------------------------------------------------------------------
// Class for export from html to doc (MS Word)
// Created by info from https://habr.com/post/168977/
//------------------------------------------------------------------------------------------------------------------------------------------------------
class ExportToWord {

	private static $mimeDocSeparator = 'doc_file_separator';	// any text for mime separation
	private static $srcEncoding = 'UTF-8';						// source project encoding
	private static $imgRef = '/img/';							// path to the images directory in 'ref' format
	private static $imgDir = 'c:/myproj/img/';					// path to the images directory in absolute format

	public static function htmlToDoc($srcHtml, $srcCss, $filePath = NULL, $filePathEncoding = 'UTF-8') {
		// Converting from html to *.doc into the file $filePath
		$entities = static::htmlEntities($srcHtml);	// Get image data
		$body = '';
		$body .= static::headDoc($srcCss);
		$body .= static::bodyDoc($entities['html']);
		$body .= $entities['imagesData'];
		$body .= static::fileList($entities);
		$body .= '--'.static::$mimeDocSeparator.'--';
		$body = mb_ereg_replace('\\t+', '', $body);		// remove tabs
		// Save to file
		if($filePath) {
			if($filePathEncoding != static::$srcEncoding) $filePath = iconv(static::$srcEncoding, $filePathEncoding, $filePath);
			file_put_contents($filePath, $body);
		}
		return $body;
	}

	private static function headDoc($css) {
		// *.doc header
		$head = 'MIME-Version: 1.0
			Content-Type: multipart/related; boundary="'.static::$mimeDocSeparator.'"

			--'.static::$mimeDocSeparator.'
			Content-Transfer-Encoding: quoted-printable
			Content-Type: text/html; charset="UTF-8"

			<html xmlns:o=3D"urn:schemas-microsoft-com:office:office"
			xmlns:w=3D"urn:schemas-microsoft-com:office:word"
			xmlns=3D"http://www.w3.org/TR/REC-html40">
			<head>
			<meta http-equiv=3DContent-Type content=3D"text/html; charset=3DUTF-8">
			<meta name=3DProgId content=3DWord.Document>
			<meta name=3DGenerator content=3D"Microsoft Word 11">
			<meta name=3DOriginator content=3D"Microsoft Word 11">
			<link rel=3DFile-List href=3D"filelist.xml">
			<!--[if gte mso 9]><xml>
				<w:WordDocument>
				<w:View>Print</w:View>
				<w:GrammarState>Clean</w:GrammarState>
				<w:ValidateAgainstSchemas/>
				<w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
				<w:IgnoreMixedContent>false</w:IgnoreMixedContent>
				<w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
				<w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
				</w:WordDocument>
			</xml><![endif]--><!--[if gte mso 9]><xml>
				<w:LatentStyles DefLockedState=3D"false" LatentStyleCount=3D"156">
				</w:LatentStyles>
			</xml><![endif]-->
			<style>
			<!--
				/* Style Definitions */
				p.MsoNormal, li.MsoNormal, div.MsoNormal
				{mso-style-parent:"";
				margin:0cm;
				margin-bottom:.0001pt;
				mso-pagination:widow-orphan;
				font-size:12.0pt;
				font-family:"Tahoma";
				mso-fareast-font-family:"Tahoma";}
			@page Section1
				{size:595.3pt 841.9pt;
				margin:1.5cm 1.5cm 1.5cm 1.5cm;
				mso-header-margin:35.4pt;
				mso-footer-margin:35.4pt;
				mso-paper-source:0;}
			div.Section1
				{page:Section1;}
			-->
			</style>
			<!--[if gte mso 10]>
			<style>
				/* Style Definitions */
				table.MsoNormalTable
				{mso-style-name:"\041E\0431\044B\0447\043D\0430\044F \0442\0430\0431\043B\=0438\0446\0430";
				mso-tstyle-rowband-size:0;
				mso-tstyle-colband-size:0;
				mso-style-noshow:yes;
				mso-style-parent:"";
				mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
				mso-para-margin:0cm;
				mso-para-margin-bottom:.0001pt;
				mso-pagination:widow-orphan;
				font-size:10.0pt;
				font-family:"Tahoma";
				mso-ansi-language:#0400;
				mso-fareast-language:#0400;
				mso-bidi-language:#0400;
				width:100%;
			}

			td.br1{
				border:1px solid black;
			}
			'.$css.'
			</style>
			<![endif]-->
			</head>';
		return $head;
	}

	private static function htmlEntities($html) {
		// Get images from html and receiving their data for embeding
		$imagesData = '';
		preg_match_all('/<img\s*(.*?)\s*src\s*=\s*"(.+?)"(.*?)>/u', $html, $matches);
		$i = 0;
		foreach($matches[0] as $imgTag) {
			$filePath = mb_ereg_replace(static::$imgRef, static::$imgDir, $matches[2][$i]);
			$imageExt = pathinfo($matches[2][$i])['extension'];
			$imageName = 'images'.$i.'.'.$imageExt;
			$imagesNames .=  '<o:File HRef=3D"'.$imageName.'"/>';
			$imageData = chunk_split(base64_encode(file_get_contents($filePath)));
			$imagesData .= '
				--'.static::$mimeDocSeparator.'
				Content-Location: images/'.$imageName.'
				Content-Transfer-Encoding: base64
				Content-Type: image/'.$imageExt.'

				'.$imageData.'
				';
			$imageDesc = '
				<v:imagedata src="images/'.$imageName.'" o:href=""/>
				</v:shape><![endif]--><![if !vml]><span style="mso-ignore:vglayout"><img border=3D0 src="images/'.$imageName.'"
				alt=3DHaut v:shapes="_x0000_i1057" '.$matches[2][$i].' '.$matches[3][$i].'></span><![endif]>';
			$html = mb_ereg_replace($imgTag, $imageDesc, $html);
			$i++;
		}
		$html = preg_replace('/=/u', '=3D', $html);
		return ['html' => $html, 'imagesNames' => $imagesNames, 'imagesData' => $imagesData];
	}

	private static function bodyDoc($body) {
		// *.doc body
		$body = '<body><div class=3D"Section1">'.$body.'</div>
			</body>
			</html>';
		return $body;
	}

	private static function footerDoc() {
		// *.doc footer
		$footer = '</div>
			</body>
			</html>
			';
		return $footer;
	}

	private static function fileList($entities) {
		// XML for embeding images
		$fileList = '
			--'.static::$mimeDocSeparator.'
			Content-Location: filelist.xml
			Content-Transfer-Encoding: quoted-printable
			Content-Type: text/xml; charset="utf-8"

			<xml xmlns:o=3D"urn:schemas-microsoft-com:office:office">
			<o:MainFile HRef=3D"../doc.doc"/>
			'.$entities['imagesNames'].'
			<o:File HRef=3D"filelist.xml"/>
			</xml>
			';
		return $fileList;
	}
}
//------------------------------------------------------------------------------------------------------------------------------------------------------
