<?php
require 'Parsedown.php';
require_once 'vendor/autoload.php';
include_once 'phpodt/phpodt.php';

// $type = Parsedown::TYPE_HTML;
// $type = Parsedown::TYPE_ODT;
$type = Parsedown::TYPE_DOCX;

/*
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$phpWord->addNumberingStyle('list', ['type' => 'multilevel', 'levels' => [['format' => 'bullet', 'text' => '-', 'left' => 520, 'hanging' => 200, 'tabPos' => 520], ['format' => 'bullet', 'text' => '-', 'left' => 880, 'hanging' => 200, 'tabPos' => 880]]]);
$section = $phpWord->addSection();
$textRun = $section->addListItemRun(0, 'list');
$textRun->addText('level 0');
$textRun = $section->addListItemRun(0, 'list');
$textRun->addText('level 0');
$textRun = $section->addListItemRun(1, 'list');
$textRun->addText('level 1');
$textRun = $section->addListItemRun(1, 'list');
$textRun->addText('level 1');
$textRun = $section->addListItemRun(0, 'list');
$textRun->addText('level 0');
$textRun = $section->addListItemRun(0, 'list');
$textRun->addText('level 0');
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('output/00000_test.docx');
*/

/*
$odt = ODT::getInstance();

$listStyle = new ListStyle('dash');
$listStyle->setBulletLevel(1, StyleConstants::RIGHT_ARROW);
$listStyle->setNumberLevel(2, new NumberFormat('', '', 'I'));
$listStyle->setNumberLevel(3, new NumberFormat('', '', '1'));

$list = new ODTList();
$list->setStyle($listStyle);
$list->addItem('TEST 1');

$sublist = new ODTList();
$sublist->addItem('TEST 1.1');

$subsublist = new ODTList();
$subsublist->addItem('TEST 1.1.1');
$subsublist->addItem('TEST 1.1.2');

$sublist->addSubList($subsublist);
$sublist->addItem('TEST 1.2');

$list->addSubList($sublist);
$list->addItem('TEST 2');
*/

// multiple_paragraphs_with_styling.mdd
// headings_and_images.mdd
// lists.mdd
// test.mdd
$inputFileFull = 'lists.mdd';
$inputFile = explode('.', $inputFileFull)[0];
$text = file_get_contents('input/'.$inputFileFull, FILE_USE_INCLUDE_PATH);
// $text = '_Hello_ world 3!';

$parsedown = new Parsedown($type);
// $tree = $parsedown->getTree($text);
// $parsedown->getHtmlFromTree($tree);
// $output = $parsedown->getDocxFromTree($tree);

$output = $parsedown->text($text);

$filename = "output/".time()."_$inputFile.$type";

if ($type == Parsedown::TYPE_ODT)
{
    ODT::getInstance()->output($filename);
}
elseif ($type == Parsedown::TYPE_DOCX)
{
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($parsedown->getPhpWord(), 'Word2007');
    $objWriter->save($filename);
}
else
{
    file_put_contents($filename, $output);
}

