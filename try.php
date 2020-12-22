<?php
require 'Parsedown.php';
require 'variableconverter/VariableConverter.php';
require_once 'vendor/autoload.php';
include_once 'phpodt/phpodt.php';

// ~~~~~~~~~~~~~ DB BEGIN ~~~~~~~~~~~~~

$idPredmet = 4780; // IIVK 1/20 - Slavica Periz
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
    $stmt = $pdo->query('SELECT k.*, m.naziv AS mesto FROM kancelarija k LEFT JOIN mesto m ON k.id_mesto = m.id ORDER BY k.id LIMIT 1');
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

$getPoverioci = function() {
    $povs = [];
    $povs[] = ['ime' => 'ЈКП "ПАРКИНГ СЕРВИС" НОВИ САД', 'adresa' => 'Нови Сад, ул. Филипа Вишњића бр. 47', 'maticni' => 'МБ 08831149, ПИБ 103635323', 'advokat' => 'чији је пуномоћник адв. Звездан Живанов, Нови Сад, Владике Платона 8/2'];
    $povs[] = ['ime' => 'Јустин Лугумерски', 'adresa' => 'Сомбор, ул. Чонопљански пут бб', 'maticni' => 'ЈМБГ 0211958810068', 'advokat' => 'чији је пуномоћник адв. Радован Зиндовић, Сомбор, Доситеја Обрадовића 15'];
    return $povs;
};

$poveriociParamFns = [function (&$poverioci, array $propsToShowWithEm) {
    if (!empty($propsToShowWithEm)) {
        // propsToShowWithEm can contain MD emphasis so need to extract that
        $propsToShow = [];
        $propsEms = [];
        $pd = new Parsedown();
        foreach ($propsToShowWithEm as $p) {
            $emMap = [];
            $key = $pd->extractEmphasis($p, $emMap);
            $propsToShow[] = $key;
            $propsEms[$key] = $emMap;
        }

        $properties = ['ime', 'adresa', 'maticni', 'advokat'];
        $propsToHide = array_diff($properties, $propsToShow);
        foreach ($poverioci as $i => &$pov) {
            foreach ($properties as $prop) {
                $pov["print_$prop"] = !in_array($prop, $propsToHide);
                $pov["ems_$prop"] = isset($propsEms[$prop]) ? $propsEms[$prop] : [];
            }
        }
    }
}];

$handlePoverioci = function($poverioci) {
    $pd = new Parsedown();
    $properties = ['ime', 'adresa', 'maticni', 'advokat'];
    $text = '';
    $num = count($poverioci);
    foreach ($poverioci as $i => $pov) {
        $p = false;
        foreach ($properties as $prop) {
            if ((!isset($pov["print_$prop"]) || $pov["print_$prop"]) && $pov[$prop]) {
                $text .= ($p ? ', ' : '');
                $text .= $pd->addEmphasisToSourceBasedOnMap($pov[$prop], isset($pov["ems_$prop"]) ? $pov["ems_$prop"] : []);
                $p = true;
            }
        }
        if ($i != $num - 1) {
            $text .= $p ? ', ' : '';
        }
    }
    return ['text' => $text, 'markdown' => true];
};

// ~~~~~~~~~~~~~ DB END ~~~~~~~~~~~~~

$varConverter = new VariableConverter();
$varConverter->registerVariable('izvrsitelj', $getIzvrsitelj);
$varConverter->registerVariable('oznaka', $getOznaka);
$varConverter->registerVariable('kancelarija', $getKancelarija, ['naziv', 'adresa', 'mesto', 'tel1', 'tel2', 'tel3', 'pib', 'maticni_broj']);
$varConverter->registerVariable('identifikacioni_broj', $getIdentifikacioniBroj);
$varConverter->registerVariable('datum_zop', $getDatumZop);
$varConverter->registerCollectionVariable('poverioci', $getPoverioci, $handlePoverioci, $poveriociParamFns);

/*

var_dump('keys:', $varConverter->getKeys());
// var_dump('values:', $varConverter->getValues());

echo $varConverter->evaluate('izvrsitelj')."\n";
echo $varConverter->evaluate('oznaka')."\n";
echo $varConverter->evaluate('identifikacioni_broj')."\n";
echo $varConverter->evaluate('kancelarija.naziv')."\n";
echo $varConverter->evaluate('datum_zop')."\n";
echo $varConverter->evaluate('poverioci', ['print' => ['ime', 'adresa']])."\n";

// var_dump('values:', $varConverter->getValues());

*/

// $type = Parsedown::TYPE_HTML;
$type = Parsedown::TYPE_ODT;
// $type = Parsedown::TYPE_DOCX;

//$phpWord = new \PhpOffice\PhpWord\PhpWord();
//$odt = ODT::getInstance();

// multiple_paragraphs_with_styling.mdd
// headings_and_images.mdd
// lists.mdd
// escaping_special_characters.mdd
// simple_variables_with_params.mdd
// test.mdd
$inputFileFull = 'simple_variables_with_params.mdd';
$inputFile = explode('.', $inputFileFull)[0];
$text = file_get_contents('input/'.$inputFileFull, FILE_USE_INCLUDE_PATH);
// $text = '**_ЈКП "ПАРКИНГ СЕРВИС" НОВИ САД_**, Нови Сад, ул. Филипа Вишњића бр. 47, **_ЈКП "ПАРКИНГ СЕРВИС" НИШ_**, НИШ, ул. Генерала Милојка Лешјанина';


$parsedown = new Parsedown($type);
$parsedown->setVarConverter($varConverter);

// $tree = $parsedown->getTree($text);
// var_dump($parsedown->lineElements($tree[0]['handler']['argument']));
// var_dump($parsedown->getHtmlFromTree($tree));
// $output = $parsedown->getDocxFromTree($tree);

/*
$map = [];
var_dump($parsedown->extractEmphasis('_***Hello***_', $map), $map);
*/

$output = $parsedown->text($text);
echo $output;

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
