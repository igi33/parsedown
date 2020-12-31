<?php
require 'Parsedown.php';
require 'variableconverter/VariableConverter.php';
require_once 'vendor/autoload.php';
include_once 'phpodt/phpodt.php';

// ~~~~~~~~~~~~~ DB BEGIN ~~~~~~~~~~~~~

$idPredmet = 4818; // IIVK 25/20 - Slavica Periz
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

$textcaseParamFn = function(&$text, $case) {
    $caseLc = strtolower($case);
    if ($caseLc == 'upper') {
        $text = mb_strtoupper($text, 'UTF-8');
    } elseif ($caseLc == 'lower') {
        $text = mb_strtolower($text, 'UTF-8');
    }
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

$getPredmet = function() use ($pdo, $idPredmet) {
    $stmt = $pdo->prepare('SELECT * FROM predmet WHERE id = ?');
    $stmt->execute([$idPredmet]);
    $predmet = $stmt->fetch();
    return $predmet;
};

$getDatumZop = function() use ($pdo, $idPredmet) {
    $stmt = $pdo->prepare('SELECT pa.datum_donosenja FROM predmet_akti pa WHERE pa.id_predmet = ? AND pa.id_akt = ? ORDER BY datum_donosenja DESC LIMIT 1');
    $stmt->execute([$idPredmet, 1]);
    $zop = $stmt->fetch();
    return $zop['datum_donosenja'];
};

$getPoverioci = function() {
    $povs = [];
    $povs[] = ['ime' => 'ЈКП "ПАРКИНГ СЕРВИС" НОВИ САД', 'adresa' => 'Нови Сад, ул. Филипа Вишњића бр. 47', 'maticni' => 'МБ 08831149, ПИБ 103635323', 'jedinicaRacun' => '', 'racun' => 'број рачуна 205-0000000128201-90 који се води код банке КОМЕРЦИЈАЛНА БАНКА А.Д. БЕОГРАД', 'advokat' => 'чији је пуномоћник адв. Звездан Живанов, Нови Сад, Владике Платона 8/2'];
    return $povs;
};

$getDuznici = function() {
    $duz = [];
    $duz[] = ['ime' => 'Немања Радмановић', 'adresa' => 'Црвенка, ул. Лењинова бр. 74', 'maticni' => 'ЈМБГ 2602998820194', 'jedinicaRacun' => '', 'racun' => '', 'advokat' => ''];
    return $duz;
};

$getRedovni = function() {
    return ['broj_racuna' => '340-11412509-08', 'naziv_banke' => 'ERSTE BANK А.Д. НОВИ САД'];
};

$getIznosPredujma = function() {
    return '5479.20';
};

$getPozivNaBroj = function() {
    $oznaka = 'предмета ИИВК 25/20';
    $pozivNaBroj = $oznaka;
    return $pozivNaBroj;
};

$getIspravu = function() {
    return 'веродостојну';
};

$datumParamFn = function (&$datum, $format) {
    if ($format) {
        $datum = date($format, strtotime($datum));
    }
};

$cenaParamFn = function (&$iznos, $tip) {
    if ($tip == 'iznos') {
        $iznos = number_format($iznos, 2, ',', '.');
    }
};

$strankaFormatParamFn = function (&$poverioci, array $propsToShowWithEm) {
    if (!empty($propsToShowWithEm)) {
        $propsToShow = [];
        $propsEms = [];
        $pd = new Parsedown();
        foreach ($propsToShowWithEm as $p) {
            $emMap = [];
            $key = $pd->extractEmphasisFromSource($p, $emMap);
            $propsToShow[] = $key;
            $propsEms[$key] = $emMap;
        }

        $properties = ['ime', 'adresa', 'maticni', 'jedinicaRacun', 'racun', 'advokat'];
        $propsToHide = array_diff($properties, $propsToShow);
        foreach ($poverioci as $i => &$pov) {
            foreach ($properties as $prop) {
                $pov["print_$prop"] = !in_array($prop, $propsToHide);
                $pov["ems_$prop"] = isset($propsEms[$prop]) ? $propsEms[$prop] : [];
            }
        }
    }
};

$handleStranke = function(array $poverioci) {
    $pd = new Parsedown();
    $properties = ['ime', 'adresa', 'maticni', 'jedinicaRacun', 'racun', 'advokat'];
    $text = '';
    $num = count($poverioci);
    foreach ($poverioci as $i => $pov) {
        $p = false;
        foreach ($properties as $prop) {
            if ((!isset($pov["print_$prop"]) || $pov["print_$prop"]) && $pov[$prop]) {
                $text .= ($p ? ', ' : '');
                $text .= $pd->insertEmphasisToSource($pov[$prop], isset($pov["ems_$prop"]) ? $pov["ems_$prop"] : []);
                $p = true;
            }
        }
        if ($i != $num - 1) {
            $text .= $p ? ', ' : '';
        }
    }
    return ['text' => $text, 'markdown' => true];
};

$getStavke = function() {
    $stavke = [];
    $stavke[] = [
        'ukupno_sa_pdv' => '1728.000',
        'ukupno_bez_pdv' => '1440.00',
        'kolicina' => '1',
        'kolicinaIskorisceno' => '1',
        'cena_jm' => '1440.00',
        'id_trosak' => '54',
        'na_ime_stavka' => null,
        'pdv' => '20',
        'zaduzenje' => '1728.00',
        'na_ime' => 'накнаде за припрему, вођење и архивирање предмета',
        'naziv_troska' => 'Накнада за припрему, вођење и архивирање предмета',
        'id_jedinica_mere' => null,
        'clan_zakona' => '9',
        'tarifni_broj' => '1',
        'stav' => null,
        'tacka' => null,
        'jedinica_mere' => null,
        'id_predmet' => '4818',
        'id_tip_predujma' => '1',
        'naziv_tipa_predujma' => 'Основни предујам',
    ];
    $stavke[] = [
      'ukupno_sa_pdv' => '194.400',
      'ukupno_bez_pdv' => '162.00',
      'kolicina' => '3',
      'kolicinaIskorisceno' => '3',
      'cena_jm' => '54.00',
      'id_trosak' => '56',
      'na_ime_stavka' => 'стварних трошкова извршног поступка',
      'pdv' => '20',
      'zaduzenje' => '194.40',
      'na_ime' => 'трошкова по члану 17.',
      'naziv_troska' => 'Трошкови по члану 17.',
      'id_jedinica_mere' => '3',
      'clan_zakona' => '17',
      'tarifni_broj' => null,
      'stav' => null,
      'tacka' => null,
      'jedinica_mere' => 'комад/а',
      'id_predmet' => '4818',
      'id_tip_predujma' => '1',
      'naziv_tipa_predujma' => 'Основни предујам',
    ];
    $stavke[] = [
        'ukupno_sa_pdv' => '2520.000',
        'ukupno_bez_pdv' => '2100.00',
        'kolicina' => '7',
        'kolicinaIskorisceno' => '7',
        'cena_jm' => '300.00',
        'id_trosak' => '62',
        'na_ime_stavka' => null,
        'pdv' => '20',
        'zaduzenje' => '2520.00',
        'na_ime' => 'накнаде за достављање поштом странкама, учесницима у поступку и суду',
        'naziv_troska' => 'Достављање поштом странкама, учесницима у поступку и суду',
        'id_jedinica_mere' => '3',
        'clan_zakona' => '18',
        'tarifni_broj' => '2',
        'stav' => null,
        'tacka' => null,
        'jedinica_mere' => 'комад/а',
        'id_predmet' => '4818',
        'id_tip_predujma' => '1',
        'naziv_tipa_predujma' => 'Основни предујам',
    ];
    $stavke[] = [
        'ukupno_sa_pdv' => '345.600',
        'ukupno_bez_pdv' => '288.00',
        'kolicina' => '1',
        'kolicinaIskorisceno' => '1',
        'cena_jm' => '288.00',
        'id_trosak' => '75',
        'na_ime_stavka' => null,
        'pdv' => '20',
        'zaduzenje' => '345.60',
        'na_ime' => 'накнаде за састављање решења о извршењу',
        'naziv_troska' => 'Састављање решења о извршењу',
        'id_jedinica_mere' => '3',
        'clan_zakona' => '18',
        'tarifni_broj' => '2',
        'stav' => null,
        'tacka' => null,
        'jedinica_mere' => 'комад/а',
        'id_predmet' => '4818',
        'id_tip_predujma' => '1',
        'naziv_tipa_predujma' => 'Основни предујам',
    ];
    $stavke[] = [
        'ukupno_sa_pdv' => '345.600',
        'ukupno_bez_pdv' => '288.00',
        'kolicina' => '1',
        'kolicinaIskorisceno' => '1',
        'cena_jm' => '288.00',
        'id_trosak' => '80',
        'na_ime_stavka' => null,
        'pdv' => '20',
        'zaduzenje' => '345.60',
        'na_ime' => 'накнаде за састављање решења о трошковима поступка',
        'naziv_troska' => 'Састављање решења о трошковима поступка',
        'id_jedinica_mere' => '3',
        'clan_zakona' => '18',
        'tarifni_broj' => '2',
        'stav' => null,
        'tacka' => null,
        'jedinica_mere' => 'комад/а',
        'id_predmet' => '4818',
        'id_tip_predujma' => '1',
        'naziv_tipa_predujma' => 'Основни предујам',
    ];
    $stavke[] = [
      'ukupno_sa_pdv' => '345.600',
      'ukupno_bez_pdv' => '288.00',
      'kolicina' => '1',
      'kolicinaIskorisceno' => '1',
      'cena_jm' => '288.00',
      'id_trosak' => '82',
      'na_ime_stavka' => null,
      'pdv' => '20',
      'zaduzenje' => '345.60',
      'na_ime' => 'накнаде за састављање закључка о предујму',
      'naziv_troska' => 'Састављање закључка о предујму',
      'id_jedinica_mere' => '3',
      'clan_zakona' => '18',
      'tarifni_broj' => '2',
      'stav' => null,
      'tacka' => null,
      'jedinica_mere' => 'комад/а',
      'id_predmet' => '4818',
      'id_tip_predujma' => '1',
      'naziv_tipa_predujma' => 'Основни предујам',
    ];
    return $stavke;
};

$handleStavke = function(array $stavke) {
    // на име накнаде за припрему, вођење и архивирање предмета износ од 1.728,00 динара (1.440,00 дин. + 288,00 дин на име 20% ПДВ-а),
    $num = count($stavke);
    $text = '';
    foreach ($stavke as $i => $si)
    {
        $stavkaSaPdv = $si['ukupno_sa_pdv'];
        $stavkaBezPdv = $si['ukupno_bez_pdv'];
        $text .= 'на име '.($si['na_ime_stavka'] ? $si['na_ime_stavka'] : $si['na_ime']);
        $text .= ' износ од ' . number_format($stavkaSaPdv, 2, ',', '.') . ' динара';
        if ($stavkaSaPdv != $stavkaBezPdv) {
            $text .= ' ('.number_format($stavkaBezPdv, 2, ',', '.').' дин. + '.number_format($stavkaSaPdv-$stavkaBezPdv, 2, ',', '.').' дин на име '.$si['pdv'].'% ПДВ-а)';
        }
        $text .= $num-1 == $i ? '' : ', ';
    }
    return $text;
};

$handleStavkeList = function(array $stavke) {
    // на име накнаде за припрему, вођење и архивирање предмета износ од 1.728,00 динара (1.440,00 дин. + 288,00 дин на име 20% ПДВ-а),
    $num = count($stavke);
    $text = '';
    foreach ($stavke as $i => $si)
    {
        $text .= '- ';
        $stavkaSaPdv = $si['ukupno_sa_pdv'];
        $stavkaBezPdv = $si['ukupno_bez_pdv'];
        $text .= 'на име '.($si['na_ime_stavka'] ? $si['na_ime_stavka'] : $si['na_ime']);
        $text .= ' износ од ' . number_format($stavkaSaPdv, 2, ',', '.') . ' динара';
        if ($stavkaSaPdv != $stavkaBezPdv) {
            $text .= ' ('.number_format($stavkaBezPdv, 2, ',', '.').' дин. + '.number_format($stavkaSaPdv-$stavkaBezPdv, 2, ',', '.').' дин на име '.$si['pdv'].'% ПДВ-а)';
        }
        $text .= PHP_EOL;
    }
    return ['text' => $text, 'markdown' => true];
};

// ~~~~~~~~~~~~~ DB END ~~~~~~~~~~~~~

$varConverter = new VariableConverter();
$varConverter->registerVariable('izvrsitelj', $getIzvrsitelj, [$textcaseParamFn]);
$varConverter->registerVariable('oznaka', $getOznaka);
$varConverter->registerVariable('kancelarija', $getKancelarija, [], ['naziv', 'adresa', 'mesto', 'tel1', 'tel2', 'tel3', 'pib', 'maticni_broj']);
$varConverter->registerVariable('predmet', $getPredmet, [$datumParamFn], ['identifikacioni_broj', 'datum_prijema', 'broj_izvrsne_verodostojne_isprave']);
$varConverter->registerVariable('datumZop', $getDatumZop, [$datumParamFn]);
$varConverter->registerCollectionVariable('poverioci', $getPoverioci, $handleStranke, [$strankaFormatParamFn]);
$varConverter->registerCollectionVariable('duznici', $getDuznici, $handleStranke, [$strankaFormatParamFn]);
$varConverter->registerVariable('iznosPredujma', $getIznosPredujma, [$cenaParamFn]);
$varConverter->registerVariable('redovni', $getRedovni, [], ['broj_racuna', 'naziv_banke']);
$varConverter->registerVariable('pozivNaBroj', $getPozivNaBroj);
$varConverter->registerVariable('ispravu', $getIspravu);
$varConverter->registerCollectionVariable('stavkePredujmaInline', $getStavke, $handleStavke);
$varConverter->registerCollectionVariable('stavkePredujmaList', $getStavke, $handleStavkeList);

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
// zop_items_inline.mdd
// zop_items_list.mdd
// tables.mdd
// test.mdd
$inputFileFull = 'test.mdd';
$inputFile = explode('.', $inputFileFull)[0];
$text = file_get_contents('input/'.$inputFileFull, FILE_USE_INCLUDE_PATH);
// $text = '**_ЈКП "ПАРКИНГ СЕРВИС" НОВИ САД_**, Нови Сад, ул. Филипа Вишњића бр. 47, **_ЈКП "ПАРКИНГ СЕРВИС" НИШ_**, НИШ, ул. Генерала Милојка Лешјанина';


$parsedown = new Parsedown($type);
$parsedown->setVarConverter($varConverter);

// $tree = $parsedown->getParseTree($text);
// var_dump($parsedown->lineElements($tree[0]['handler']['argument']));

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
