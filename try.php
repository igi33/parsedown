<?php
require 'Parsedown.php';
require 'variableconverter/VariableConverter.php';
require_once 'vendor/autoload.php';
include_once 'phpodt/phpodt.php';

$idPredmet = 4780; // IIVK 1/20 - Slavica Periz

// ~~~~~~~~~~~~~ DB BEGIN ~~~~~~~~~~~~~

$host = '127.0.0.1';
$db   = 'slavica_periz_so_izv';
$user = 'root';
$pass = '';
$charset = 'utf8';

$dsn = "mysql:host=$host;dbname=$db;charset=$charset";
$options = [
    PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
    PDO::ATTR_EMULATE_PREPARES   => false,
];
try {
     $pdo = new PDO($dsn, $user, $pass, $options);
} catch (\PDOException $e) {
     throw new \PDOException($e->getMessage(), (int)$e->getCode());
}

$getIzvrsitelj = function() use ($pdo) {
    $stmt = $pdo->query('SELECT * FROM izvrsitelj ORDER BY id LIMIT 1');
    $izv = $stmt->fetch();
    return $izv['ime'].' '.$izv['prezime'];
};

$getKancelarija = function() use ($pdo) {
    $stmt = $pdo->query('SELECT * FROM kancelarija ORDER BY id LIMIT 1');
    $kanc = $stmt->fetch();
    return $kanc;
};

$getOznaka = function() use ($pdo, $idPredmet) {
    $stmt = $pdo->prepare('SELECT broj_predmeta, godina, u.upisnik FROM predmet p LEFT JOIN upisnik u ON p.id_upisnik = u.id WHERE p.id = ?');
    $stmt->execute([$idPredmet]);
    $predmet = $stmt->fetch();
    return $predmet['upisnik'] . ' ' . $predmet['broj_predmeta'] . '/' . substr($predmet['godina'], 2);
};

$getIdentifikacioniBroj = function() use ($pdo, $idPredmet) {
    $stmt = $pdo->prepare('SELECT identifikacioni_broj FROM predmet WHERE id = ?');
    $stmt->execute([$idPredmet]);
    $predmet = $stmt->fetch();
    return $predmet['identifikacioni_broj'];
};

$getDatumZop = function() use ($pdo, $idPredmet) {
    $stmt = $pdo->prepare('SELECT pa.datum_donosenja FROM predmet_akti pa WHERE pa.id_predmet = ? AND pa.id_akt = ? ORDER BY datum_donosenja DESC LIMIT 1');
    $stmt->execute([$idPredmet, 1]);
    $zop = $stmt->fetch();
    return date('d.m.Y.', strtotime($zop['datum_donosenja']));
};

// ~~~~~~~~~~~~~ DB END ~~~~~~~~~~~~~


$varConverter = new VariableConverter();
$varConverter->registerVariable('izvrsitelj', $getIzvrsitelj);
$varConverter->registerVariable('oznaka', $getOznaka);
$varConverter->registerVariable('kancelarija', $getKancelarija);
$varConverter->registerVariable('identifikacioni_broj', $getIdentifikacioniBroj);
$varConverter->registerVariable('datum_zop', $getDatumZop);

var_dump('keys:', $varConverter->getKeys());
// var_dump('values:', $varConverter->getValues());

echo $varConverter->evaluate('izvrsitelj')."\n";
echo $varConverter->evaluate('oznaka')."\n";
echo $varConverter->evaluate('identifikacioni_broj')."\n";
echo $varConverter->evaluate('kancelarija')['naziv']."\n";
echo $varConverter->evaluate('datum_zop')."\n";

// var_dump('values:', $varConverter->getValues());

/*

// $type = Parsedown::TYPE_HTML;
$type = Parsedown::TYPE_ODT;
// $type = Parsedown::TYPE_DOCX;

//$phpWord = new \PhpOffice\PhpWord\PhpWord();
//$odt = ODT::getInstance();

// multiple_paragraphs_with_styling.mdd
// headings_and_images.mdd
// lists.mdd
// escaping_special_characters.mdd
// test.mdd
$inputFileFull = 'test.mdd';
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
*/
