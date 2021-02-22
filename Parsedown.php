<?php
include_once 'phpodt/phpodt.php';

#
#
# Parsedown
# http://parsedown.org
#
# (c) Emanuil Rusev
# http://erusev.com
#
# For the full license information, view the LICENSE file that was distributed
# with this source code.
#
#

class Parsedown
{
    # ~

    const version = '1.8.0-beta-7';

    // supported output type constants
    const TYPE_HTML = 'html';
    const TYPE_DOCX = 'docx';
    const TYPE_ODT = 'odt';

    # ~

    protected $outputMode; // html, docx, odt
    protected $varConverter; // instance of VariableConverter

    // document properties
    protected $fontName;
    protected $fontSize;
    protected $marginTop; // in cm
    protected $marginBottom; // in cm
    protected $marginLeft; // in cm
    protected $marginRight; // in cm
    
    // shared properties
    protected $options; // holds the current format style state in a map, e.g. bold, italic, underline
    protected $paragraphDepth; // based on this property a paragraph style is chosen
    protected $listDepth;
    protected $listIsRoman; // boolean property: true => roman lists, false => decimal lists
    protected $listIsBullets; // boolean property: true => bullet lists, false => hyphen lists
    protected $table;
    protected $row;
    protected $isTh; // true if in thead, false if in tbody
    protected $pStyleCell; // holds the current pStyle name for table cell text
    protected $pStyle; // holds the current pStyle name for paragraph text
    protected $headingsMap;
    protected $currentColSpan; // currently only used for DOCX, determines table cell creation
    protected $addTextToCell; // currently only for DOCX, determines whether text should be added to cell
    protected $sectionIdx; // current id of section

    // properties used for keeping phpword internal state
    protected $phpWord;
    protected $section; // holds the current section object
    protected $textRun; // holds the current textRun object
    protected $listItemRun; // holds the current listItemRun object
    protected $listStyleName;
    protected $cell;
    protected $romanNumberingStyle;
    protected $decimalNumberingStyle;

    // properties used for keeping phpodt internal state
    protected $odt;
    protected $odtTextStyles;
    protected $odtParagraphStyles;
    protected $odtListStyles;
    protected $odtPageStyle;
    protected $p; // holds the current Paragraph object
    protected $list; // holds the current ODTList object
    protected $odtFinishedListItems;
    protected $tableRows;
    protected $pageBreakBefore;
    protected $odtBaseTextStyleNames;
    protected $tableCellStylesMap;

    // properties used for keeping html internal state
    protected $markup = '';

    public function __construct($mode = null)
    {
        if ($mode)
        {
            $this->setOutputMode($mode);
        }
    }

    #
    # Initialization methods
    #

    protected function initializeSharedParameters()
    {
        $this->options = array();
        $this->marginTop = 1.06;
        $this->marginBottom = 1.06;
        $this->marginRight = 1.41;
        $this->marginLeft = 1.41;
        $this->fontSize = 11;
        $this->paragraphDepth = -1;
        $this->listDepth = -1;
        $this->listIsRoman = true;
        $this->listIsBullets = false;
        $this->table = null;
        $this->row = null;
        $this->isTh = null;
        $this->pStyle = null;
        $this->pStyleCell = null;
        $this->headingsMap = array(1 => 1, 2 => 11, 3 => 12, 4 => 13);
        $this->currentColSpan = 0;
        $this->addTextToCell = null;
        $this->sectionIdx = -1;
    }

    protected function initializeOdtParameters()
    {
        $this->fontName = 'Liberation Serif';
        $this->odt = ODT::getInstance();
        $this->odtTextStyles = array();
        $this->odtParagraphStyles = array();
        $this->odtListStyles = array();
        $this->odtPageStyle = null;
        $this->odtFinishedListItems = array();
        $this->p = null;
        $this->list = null;
        $this->tableRows = null;
        $this->pageBreakBefore = false;
        $this->odtBaseTextStyleNames = array();
        $this->tableCellStylesMap = null;

        $defaultTextStyle = new TextStyle('default');
        $defaultTextStyle->setFontName($this->fontName);
        $defaultTextStyle->setFontSize($this->fontSize);
        $defaultTextStyle->setColor();
        $this->odtTextStyles[$defaultTextStyle->getStyleName()] = $defaultTextStyle;
        $this->odtBaseTextStyleNames[] = $defaultTextStyle->getStyleName();
        
        $boldTextStyle = new TextStyle('bold');
        $boldTextStyle->setBold();
        $boldTextStyle->setFontName($this->fontName);
        $boldTextStyle->setFontSize($this->fontSize);
        $boldTextStyle->setColor();
        $this->odtTextStyles[$boldTextStyle->getStyleName()] = $boldTextStyle;
        $this->odtBaseTextStyleNames[] = $boldTextStyle->getStyleName();
        
        $italicTextStyle = new TextStyle('italic');
        $italicTextStyle->setItalic();
        $italicTextStyle->setFontName($this->fontName);
        $italicTextStyle->setFontSize($this->fontSize);
        $italicTextStyle->setColor();
        $this->odtTextStyles[$italicTextStyle->getStyleName()] = $italicTextStyle;
        $this->odtBaseTextStyleNames[] = $italicTextStyle->getStyleName();

        $underlineTextStyle = new TextStyle('underline');
        $underlineTextStyle->setTextUnderline();
        $underlineTextStyle->setFontName($this->fontName);
        $underlineTextStyle->setFontSize($this->fontSize);
        $underlineTextStyle->setColor();
        $this->odtTextStyles[$underlineTextStyle->getStyleName()] = $underlineTextStyle;
        $this->odtBaseTextStyleNames[] = $underlineTextStyle->getStyleName();
        
        $boldItalicTextStyle = new TextStyle('bold_italic');
        $boldItalicTextStyle->setBold();
        $boldItalicTextStyle->setItalic();
        $boldItalicTextStyle->setFontName($this->fontName);
        $boldItalicTextStyle->setFontSize($this->fontSize);
        $boldItalicTextStyle->setColor();
        $this->odtTextStyles[$boldItalicTextStyle->getStyleName()] = $boldItalicTextStyle;
        $this->odtBaseTextStyleNames[] = $boldItalicTextStyle->getStyleName();

        $boldUnderlineTextStyle = new TextStyle('bold_underline');
        $boldUnderlineTextStyle->setBold();
        $boldUnderlineTextStyle->setTextUnderline();
        $boldUnderlineTextStyle->setFontName($this->fontName);
        $boldUnderlineTextStyle->setFontSize($this->fontSize);
        $boldUnderlineTextStyle->setColor();
        $this->odtTextStyles[$boldUnderlineTextStyle->getStyleName()] = $boldUnderlineTextStyle;
        $this->odtBaseTextStyleNames[] = $boldUnderlineTextStyle->getStyleName();

        $italicUnderlineTextStyle = new TextStyle('italic_underline');
        $italicUnderlineTextStyle->setItalic();
        $italicUnderlineTextStyle->setTextUnderline();
        $italicUnderlineTextStyle->setFontName($this->fontName);
        $italicUnderlineTextStyle->setFontSize($this->fontSize);
        $italicUnderlineTextStyle->setColor();
        $this->odtTextStyles[$italicUnderlineTextStyle->getStyleName()] = $italicUnderlineTextStyle;
        $this->odtBaseTextStyleNames[] = $italicUnderlineTextStyle->getStyleName();
        
        $boldItalicUnderlineTextStyle = new TextStyle('bold_italic_underline');
        $boldItalicUnderlineTextStyle->setBold();
        $boldItalicUnderlineTextStyle->setItalic();
        $boldItalicUnderlineTextStyle->setTextUnderline();
        $boldItalicUnderlineTextStyle->setFontName($this->fontName);
        $boldItalicUnderlineTextStyle->setFontSize($this->fontSize);
        $boldItalicUnderlineTextStyle->setColor();
        $this->odtTextStyles[$boldItalicUnderlineTextStyle->getStyleName()] = $boldItalicUnderlineTextStyle;
        $this->odtBaseTextStyleNames[] = $boldItalicUnderlineTextStyle->getStyleName();

        $heading1 = new ParagraphStyle('h1');
        $heading1->setTextAlign(StyleConstants::CENTER);
        $heading1->setVerticalMargin('0.2cm', '0.2cm');
        $this->odtParagraphStyles[$heading1->getStyleName()] = $heading1;

        $heading2 = new ParagraphStyle('h11');
        $heading2->setTextAlign(StyleConstants::CENTER);
        $heading2->setVerticalMargin('0.2cm', '0cm');
        $this->odtParagraphStyles[$heading2->getStyleName()] = $heading2;

        $heading3 = new ParagraphStyle('h12');
        $heading3->setTextAlign(StyleConstants::CENTER);
        $heading3->setVerticalMargin('0cm', '0cm');
        $this->odtParagraphStyles[$heading3->getStyleName()] = $heading3;

        $heading4 = new ParagraphStyle('h13');
        $heading4->setTextAlign(StyleConstants::CENTER);
        $heading4->setVerticalMargin('0cm', '0.2cm');
        $this->odtParagraphStyles[$heading4->getStyleName()] = $heading4;


        $header = new ParagraphStyle('header');
        $header->setTextAlign(StyleConstants::LEFT);
        $header->setHorizontalMargins('0.0cm', '8.82cm');
        $this->odtParagraphStyles[$header->getStyleName()] = $header;

        $nonIndented = new ParagraphStyle('non_indented');
        $nonIndented->setTextAlign(StyleConstants::JUSTIFY);
        $nonIndented->setVerticalMargin('0.2cm', '0.2cm');
        $this->odtParagraphStyles[$nonIndented->getStyleName()] = $nonIndented;
        
        $indented = new ParagraphStyle('indented');
        $indented->setTextAlign(StyleConstants::JUSTIFY);
        $indented->setTextIndent('1cm');
        $indented->setVerticalMargin('0.2cm', '0.2cm');
        $this->odtParagraphStyles[$indented->getStyleName()] = $indented;
        

        $centerAligned = new ParagraphStyle('center_aligned');
        $centerAligned->setTextAlign(StyleConstants::CENTER);
        $this->odtParagraphStyles[$centerAligned->getStyleName()] = $centerAligned;
        
        $leftAligned = new ParagraphStyle('left_aligned');
        $leftAligned->setTextAlign(StyleConstants::LEFT);
        $this->odtParagraphStyles[$leftAligned->getStyleName()] = $leftAligned;
        
        $rightAligned = new ParagraphStyle('right_aligned');
        $rightAligned->setTextAlign(StyleConstants::RIGHT);
        $this->odtParagraphStyles[$rightAligned->getStyleName()] = $rightAligned;
        
        $justify = new ParagraphStyle('justify');
        $justify->setTextAlign(StyleConstants::JUSTIFY);
        $this->odtParagraphStyles[$justify->getStyleName()] = $justify;


        $newPage = new ParagraphStyle('new_page');
        $newPage->setBreakAfter(StyleConstants::PAGE);
        $this->odtParagraphStyles[$newPage->getStyleName()] = $newPage;


        $romanList = new ListStyle('roman');
        $romanList->setNumberLevel(1, new NumberFormat('', '', 'I'), $boldTextStyle);
        $romanList->setNumberLevel(2, new NumberFormat('', '', 'I'), $boldTextStyle);
        $romanList->setNumberLevel(3, new NumberFormat('', '', 'I'), $boldTextStyle);
        $this->odtListStyles[$romanList->getStyleName()] = $romanList;

        $decimalList = new ListStyle('decimal');
        $decimalList->setNumberLevel(1, new NumberFormat('', '.', '1'), $boldTextStyle);
        $decimalList->setNumberLevel(2, new NumberFormat('', '.', '1'), $boldTextStyle);
        $decimalList->setNumberLevel(3, new NumberFormat('', '.', '1'), $boldTextStyle);
        $this->odtListStyles[$decimalList->getStyleName()] = $decimalList;
        
        $bulletList = new ListStyle('bullet');
        $bulletList->setBulletLevel(1, StyleConstants::BULLET);
        $bulletList->setBulletLevel(2, StyleConstants::BULLET);
        $bulletList->setBulletLevel(3, StyleConstants::BULLET);
        $this->odtListStyles[$bulletList->getStyleName()] = $bulletList;
        
        $hyphenList = new ListStyle('hyphen');
        $hyphenList->setBulletLevel(1, StyleConstants::HYPHEN);
        $hyphenList->setBulletLevel(2, StyleConstants::HYPHEN);
        $hyphenList->setBulletLevel(3, StyleConstants::HYPHEN);
        $this->odtListStyles[$hyphenList->getStyleName()] = $hyphenList;


        $pageStyle = new PageStyle('page');
        $this->odtPageStyle = $pageStyle;
        $this->setPhpOdtMargins();
    }

    protected function initializePhpWordParameters()
    {
        $this->fontName = 'Times New Roman';
        $this->textRun = null;
        $this->listItemRun = null;
        $this->listStyleName = null;
        $this->cell = null;
        $this->section = null;
        
        $this->romanNumberingStyle = array('type' => 'multilevel','levels' => array(array('format' => 'upperRoman', 'text' => '%1', 'left' => 720, 'hanging' => 720, 'tabPos' => 720), array('format' => 'lowerRoman', 'text' => '%2', 'left' => 1000, 'hanging' => 720, 'tabPos' => 1000)));
        $this->decimalNumberingStyle = array('type' => 'multilevel','levels' => array(array('format' => 'decimal', 'text' => '%1.', 'left' => 720, 'hanging' => 720, 'tabPos' => 720), array('format' => 'decimal', 'text' => '%2.', 'left' => 1000, 'hanging' => 720, 'tabPos' => 1000)));
        
        \PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);
        $this->phpWord = new \PhpOffice\PhpWord\PhpWord();
        $this->phpWord->setDefaultFontName($this->fontName);
        $this->phpWord->setDefaultFontSize($this->fontSize);
        $this->phpWord->setDefaultParagraphStyle(array('align' => 'left', 'spaceAfter' => 0, 'spaceBefore' => 0));
        $this->phpWord->addParagraphStyle('indented', array('spaceAfter' => 100, 'spaceBefore' => 100, 'align' => 'both', 'indentation' => array('firstLine' => 500)));
        $this->phpWord->addParagraphStyle('non_indented', array('spaceAfter' => 100, 'spaceBefore' => 100, 'align' => 'both'));
        $this->phpWord->addParagraphStyle('left_aligned', array('align' => 'left', 'spaceAfter' => 0, 'spaceBefore' => 0));
        $this->phpWord->addParagraphStyle('right_aligned', array('align' => 'right', 'spaceAfter' => 0, 'spaceBefore' => 0));
        $this->phpWord->addParagraphStyle('center_aligned', array('align' => 'center', 'spaceAfter' => 0, 'spaceBefore' => 0));
        $this->phpWord->addParagraphStyle('justify', array('align' => 'both', 'spaceAfter' => 0, 'spaceBefore' => 0));
        $this->phpWord->addParagraphStyle('header', array('spaceAfter' => 0, 'spaceBefore' => 0, 'align' => 'left', 'indentation' => array('right' => 5000)));
        $this->phpWord->addTitleStyle(1, array('bold' => true), array('spaceAfter' => 200, 'spaceBefore' => 200, 'align' => 'center'));
        $this->phpWord->addTitleStyle(11, array('bold' => true), array('spaceAfter' => 0, 'spaceBefore' => 200, 'align' => 'center'));
        $this->phpWord->addTitleStyle(12, array('bold' => true), array('spaceAfter' => 0, 'spaceBefore' => 0, 'align' => 'center'));
        $this->phpWord->addTitleStyle(13, array('bold' => true), array('spaceAfter' => 200, 'spaceBefore' => 0, 'align' => 'center'));
        $this->phpWord->addNumberingStyle('bullet', array('type' => 'multilevel', 'levels' => array(array('format' => 'bullet', 'text' => '•', 'left' => 520, 'hanging' => 200, 'tabPos' => 520), array('format' => 'bullet', 'text' => '•', 'left' => 880, 'hanging' => 200, 'tabPos' => 880), array('format' => 'bullet', 'text' => '•', 'left' => 1080, 'hanging' => 200, 'tabPos' => 1080))));
        $this->phpWord->addNumberingStyle('hyphen', array('type' => 'multilevel', 'levels' => array(array('format' => 'bullet', 'text' => '-', 'left' => 520, 'hanging' => 200, 'tabPos' => 520), array('format' => 'bullet', 'text' => '-', 'left' => 880, 'hanging' => 200, 'tabPos' => 880), array('format' => 'bullet', 'text' => '-', 'left' => 1080, 'hanging' => 200, 'tabPos' => 1080))));
    }

    #
    # PhpWord specific margin setter
    #

    protected function setPhpWordMargins()
    {
        if ($this->section)
        {
            $this->section->setStyle(array(
                'marginTop' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip($this->marginTop),
                'marginBottom' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip($this->marginBottom),
                'marginRight' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip($this->marginRight),
                'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip($this->marginLeft),
            ));
        }
    }

    #
    # PhpOdt specific margin setter
    #

    protected function setPhpOdtMargins()
    {
        if ($this->odtPageStyle)
        {
            $this->odtPageStyle->setVerticalMargin($this->marginTop.'cm', $this->marginBottom.'cm');
            $this->odtPageStyle->setHorizontalMargin($this->marginLeft.'cm', $this->marginRight.'cm');
        }
    }

    #
    # Output mode setter and getter
    #

    public function setOutputMode($mode, $initialize = true)
    {
        if (!in_array($mode, array(self::TYPE_ODT, self::TYPE_DOCX, self::TYPE_HTML)))
        {
            throw new Exception('Unknown document type');
        }

        $this->outputMode = $mode;
        
        if ($initialize)
        {
            $this->initializeSharedParameters();

            if ($mode == self::TYPE_ODT)
            {
                $this->initializeOdtParameters();
            }
            elseif ($mode == self::TYPE_DOCX)
            {
                $this->initializePhpWordParameters();
            }
        }
    }

    public function getOutputMode()
    {
        return $this->outputMode;
    }

    #
    # PhpWord instance setter and getter
    #

    public function setPhpWord(\PhpOffice\PhpWord\PhpWord $phpWord)
    {
        $this->outputMode = self::TYPE_DOCX;
        $this->phpWord = $phpWord;
    }

    public function getPhpWord()
    {
        return $this->phpWord;
    }

    #
    # VariableConverter setter and getter
    #

    public function setVarConverter(VariableConverter $varConverter)
    {
        $this->varConverter = $varConverter;
    }

    public function getVarConverter()
    {
        return $this->varConverter;
    }
    
    #
    # Get ODT page style
    #
    
    public function getOdtPageStyle()
    {
        return $this->odtPageStyle;
    }
    
    #
    # Get PhpWord section
    #
    
    public function getSection()
    {
        return $this->section;
    }

    #
    # Font name setter and getter
    #

    public function setFontName($fontName)
    {
        $this->fontName = $fontName;

        if ($this->outputMode == self::TYPE_DOCX && $this->phpWord)
        {
            $this->phpWord->setDefaultFontName($this->fontName);
        }
        elseif ($this->outputMode == self::TYPE_ODT && $this->odt)
        {
            foreach ($this->odtTextStyles as $name => $style)
            {
                $style->setFontName($this->fontName);
            }
        }
    }

    public function getFontName()
    {
        return $this->fontName;
    }

    #
    # Font size setter and getter
    #

    public function setFontSize($size)
    {
        $this->fontSize = $size;

        if ($this->outputMode == self::TYPE_DOCX && $this->phpWord)
        {
            $this->phpWord->setDefaultFontSize($this->fontSize);
        }
        elseif ($this->outputMode == self::TYPE_ODT && $this->odt)
        {
            foreach ($this->odtTextStyles as $name => $style)
            {
                $style->setFontSize($this->fontSize);
            }
        }
    }

    public function getFontSize()
    {
        return $this->fontSize;
    }

    #
    # Margin setters and getters
    #

    public function setMargins($marginTop, $marginRight, $marginBottom = null, $marginLeft = null)
    {
        $this->marginTop = $marginTop;
        $this->marginRight = $marginRight;
        $this->marginBottom = $marginBottom ? $marginBottom : $marginTop;
        $this->marginLeft = $marginLeft ? $marginLeft : $marginRight;

        if ($this->outputMode == self::TYPE_DOCX)
        {
            $this->setPhpWordMargins();
        }
        elseif ($this->outputMode == self::TYPE_ODT)
        {
            $this->setPhpOdtMargins();
        }
    }

    public function setMarginTop($marginTop)
    {
        $this->marginTop = $marginTop;

        if ($this->outputMode == self::TYPE_DOCX)
        {
            $this->setPhpWordMargins();
        }
        elseif ($this->outputMode == self::TYPE_ODT)
        {
            $this->setPhpOdtMargins();
        }
    }

    public function setMarginRight($marginRight)
    {
        $this->marginRight = $marginRight;
        
        if ($this->outputMode == self::TYPE_DOCX)
        {
            $this->setPhpWordMargins();
        }
        elseif ($this->outputMode == self::TYPE_ODT)
        {
            $this->setPhpOdtMargins();
        }
    }

    public function setMarginBottom($marginBottom)
    {
        $this->marginBottom = $marginBottom;
        
        if ($this->outputMode == self::TYPE_DOCX)
        {
            $this->setPhpWordMargins();
        }
        elseif ($this->outputMode == self::TYPE_ODT)
        {
            $this->setPhpOdtMargins();
        }
    }

    public function setMarginLeft($marginLeft)
    {
        $this->marginLeft = $marginLeft;
        
        if ($this->outputMode == self::TYPE_DOCX)
        {
            $this->setPhpWordMargins();
        }
        elseif ($this->outputMode == self::TYPE_ODT)
        {
            $this->setPhpOdtMargins();
        }
    }

    public function getMarginTop()
    {
        return $this->marginTop;
    }

    public function getMarginRight()
    {
        return $this->marginRight;
    }

    public function getMarginBottom()
    {
        return $this->marginBottom;
    }

    public function getMarginLeft()
    {
        return $this->marginLeft;
    }
    
    #
    # Re-map title (heading) id to newly created title style.
    #
    
    public function changeTitleId($level, $titleId)
    {
        if (is_int($level) && $level >= 1 && $level <= 4)
        {
            $this->headingsMap[$level] = $titleId;
        }
    }
    
    #
    # Ordered lists type: true => roman, false => decimal
    #
    
    public function setListToRoman($roman = true)
    {
        $this->listIsRoman = (bool) $roman;
    }
    
    public function isListRoman()
    {
        return $this->listIsRoman;
    }
    
    #
    # Unordered lists type: true => bullet, false => decimal
    #
    
    public function setListToBullets($bullets = true)
    {
        $this->listIsBullets = (bool) $bullets;
    }
    
    public function isListBullets()
    {
        return $this->listIsBullets;
    }

    #
    # Adds new page
    #

    public function addPage()
    {
        if ($this->outputMode == Parsedown::TYPE_ODT)
        {
            new Paragraph($this->getOdtParagraphStyle('new_page'));
        }
        elseif ($this->outputMode == Parsedown::TYPE_DOCX)
        {
            $this->section = $this->phpWord->addSection();
            $this->sectionIdx += 1;
            $this->phpWord->addNumberingStyle('roman'.$this->sectionIdx, $this->romanNumberingStyle);
            $this->phpWord->addNumberingStyle('decimal'.$this->sectionIdx, $this->decimalNumberingStyle);
            $this->setPhpWordMargins();
        }
    }

    #
    # Output methods
    #

    public function generateFile($filepath)
    {
        $filepath .= '.'.$this->outputMode;

        if ($this->outputMode == Parsedown::TYPE_ODT)
        {
            $this->odt->output($filepath);
        }
        elseif ($this->outputMode == Parsedown::TYPE_DOCX)
        {
            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->phpWord, 'Word2007');
            $objWriter->save($filepath);
        }
        else
        {
            file_put_contents($filepath, $this->markup);
        }
    }

    public function generateDownload($filename)
    {
        $filename = str_replace('/', '_', $filename).'.'.$this->outputMode;

        if ($this->outputMode == Parsedown::TYPE_ODT)
        {
            $this->odt->output($filename);
            ob_end_clean();
            header('Content-Type: application/vnd.oasis.opendocument.text');
            header('Content-type: application/force-download');
            header('Content-Disposition: attachment;filename="'.$filename.'"');
            header('Cache-Control: max-age=0'); // no cache
            $content = file_get_contents($filename);
            echo $content;
            unlink($filename);
        }
        elseif ($this->outputMode == Parsedown::TYPE_DOCX)
        {
            header('Content-Type: application/vnd.ms-word');
            header('Content-type: application/force-download');
            header('Content-Disposition: attachment;filename="'.$filename.'"');
            header('Cache-Control: max-age=0'); // no cache
            header("Content-Transfer-Encoding: binary");
            $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->phpWord, 'Word2007');
            $objWriter->save('php://output');
        }
        else
        {
            ob_end_clean();
            header('Content-Type: text/html');
            header('Content-type: application/force-download');
            header('Content-Disposition: attachment;filename="'.$filename.'"');
            header('Cache-Control: max-age=0'); // no cache
            echo $this->markup;
        }
    }

    #
    # Convenient public parse tree method
    #

    public function getParseTree($text)
    {
        return $this->textElements($text);
    }

    #
    # Process multi-line text as a section (inserts page breaks)
    #

    public function section($text)
    {
        # parse
        $Elements = $this->textElements($text);

        # convert
        if (!$this->outputMode)
        {
            throw new Exception('Output mode not set');
        }

        if ($this->outputMode == self::TYPE_ODT)
        {            
            if ($this->pageBreakBefore)
            {
                new Paragraph($this->getOdtParagraphStyle('new_page'));
            }
            else
            {
                $this->pageBreakBefore = true;
            }

            $this->elementsOdt($Elements);
        }
        elseif ($this->outputMode == self::TYPE_DOCX)
        {
            $this->section = $this->phpWord->addSection();
            $this->sectionIdx += 1;
            $this->phpWord->addNumberingStyle('roman'.$this->sectionIdx, $this->romanNumberingStyle);
            $this->phpWord->addNumberingStyle('decimal'.$this->sectionIdx, $this->decimalNumberingStyle);

            $this->setPhpWordMargins();
            $this->elementsDocx($Elements);
        }
        elseif ($this->outputMode == self::TYPE_HTML)
        {
            $this->elements($Elements);
            # trim line breaks
            $this->markup = trim($this->markup, "\n");
        }
    }

    #
    # Process multi-line text
    #

    function text($text)
    {
        # parse
        $Elements = $this->textElements($text);

        # convert
        if (!$this->outputMode)
        {
            throw new Exception('Output mode not set');
        }

        if ($this->outputMode == self::TYPE_ODT)
        {
            $this->elementsOdt($Elements);
        }
        elseif ($this->outputMode == self::TYPE_DOCX)
        {
            if ($this->section == null)
            {
                $this->section = $this->phpWord->addSection();
                $this->sectionIdx += 1;
                $this->phpWord->addNumberingStyle('roman'.$this->sectionIdx, $this->romanNumberingStyle);
                $this->phpWord->addNumberingStyle('decimal'.$this->sectionIdx, $this->decimalNumberingStyle);
                $this->setPhpWordMargins();
            }
            $this->elementsDocx($Elements);
        }
        elseif ($this->outputMode == self::TYPE_HTML)
        {
            $this->elements($Elements);
            # trim line breaks
            $this->markup = trim($this->markup, "\n");
        }
    }

    protected function textElements($text)
    {
        # make sure no definitions are set
        $this->DefinitionData = array();

        # standardize line breaks
        $text = str_replace(array("\r\n", "\r"), "\n", $text);

        # remove surrounding line breaks
        $text = trim($text, "\n");

        # split text into lines
        $lines = explode("\n", $text);

        # iterate through lines to identify blocks
        return $this->linesElements($lines);
    }

    #
    # Setters
    #

    function setBreaksEnabled($breaksEnabled)
    {
        $this->breaksEnabled = $breaksEnabled;

        return $this;
    }

    protected $breaksEnabled = true;

    function setMarkupEscaped($markupEscaped)
    {
        $this->markupEscaped = $markupEscaped;

        return $this;
    }

    protected $markupEscaped;

    function setUrlsLinked($urlsLinked)
    {
        $this->urlsLinked = $urlsLinked;

        return $this;
    }

    protected $urlsLinked = true;

    function setSafeMode($safeMode)
    {
        $this->safeMode = (bool) $safeMode;

        return $this;
    }

    protected $safeMode;

    function setStrictMode($strictMode)
    {
        $this->strictMode = (bool) $strictMode;

        return $this;
    }

    protected $strictMode;

    protected $safeLinksWhitelist = array(
        'http://',
        'https://',
        'ftp://',
        'ftps://',
        'mailto:',
        'tel:',
        'data:image/png;base64,',
        'data:image/gif;base64,',
        'data:image/jpeg;base64,',
        'irc:',
        'ircs:',
        'git:',
        'ssh:',
        'news:',
        'steam:',
    );

    #
    # Lines
    #

    protected $BlockTypes = array(
        '#' => array('Header'),
        '*' => array('List'),
        '+' => array('List'),
        '-' => array('Table', 'Rule', 'List'),
        '0' => array('List'),
        '1' => array('List'),
        '2' => array('List'),
        '3' => array('List'),
        '4' => array('List'),
        '5' => array('List'),
        '6' => array('List'),
        '7' => array('List'),
        '8' => array('List'),
        '9' => array('List'),
        ':' => array('Table'),
        '<' => array('Comment', 'Markup'),
        //'=' => array('SetextHeader'),
        '>' => array('Quote'),
        '[' => array('ParagraphStyle'),
        '_' => array('Rule'),
        // '`' => array('FencedCode'),
        '|' => array('Table'),
        // '~' => array('FencedCode'),
    );

    # ~

    protected $unmarkedBlockTypes = array(
        // 'Code', // uncomment to handle whitespace defined code blocks
    );

    #
    # Blocks
    #

    protected function lines(array $lines)
    {
        return $this->elements($this->linesElements($lines));
    }

    protected function linesElements(array $lines)
    {
        $Elements = array();
        $CurrentBlock = null;

        foreach ($lines as $line)
        {
            if (chop($line) === '')
            {
                if (isset($CurrentBlock))
                {
                    $CurrentBlock['interrupted'] = (isset($CurrentBlock['interrupted'])
                        ? $CurrentBlock['interrupted'] + 1 : 1
                    );
                }

                continue;
            }

            while (($beforeTab = strstr($line, "\t", true)) !== false)
            {
                $shortage = 4 - mb_strlen($beforeTab, 'utf-8') % 4;

                $line = $beforeTab
                    . str_repeat(' ', $shortage)
                    . substr($line, strlen($beforeTab) + 1)
                ;
            }

            $indent = strspn($line, ' ');

            $text = $indent > 0 ? substr($line, $indent) : $line;

            # ~

            $Line = array('body' => $line, 'indent' => $indent, 'text' => $text);

            # ~

            if (isset($CurrentBlock['continuable']))
            {
                $methodName = 'block' . $CurrentBlock['type'] . 'Continue';
                $Block = $this->$methodName($Line, $CurrentBlock);

                if (isset($Block))
                {
                    $CurrentBlock = $Block;

                    continue;
                }
                else
                {
                    if ($this->isBlockCompletable($CurrentBlock['type']))
                    {
                        $methodName = 'block' . $CurrentBlock['type'] . 'Complete';
                        $CurrentBlock = $this->$methodName($CurrentBlock);
                    }
                }
            }

            # ~

            $marker = $text[0];

            # ~

            $blockTypes = $this->unmarkedBlockTypes;



            if (isset($this->BlockTypes[$marker]))
            {
                foreach ($this->BlockTypes[$marker] as $blockType)
                {
                    $blockTypes []= $blockType;
                }
            }

            #
            # ~
            
            if (isset($Block['type']) && in_array($Block['type'], array('Header', 'Rule')))
            {
                $Block = null;
            }
            
            foreach ($blockTypes as $blockType)
            {
                $Block = $this->{"block$blockType"}($Line, $CurrentBlock);

                if (isset($Block))
                {
                    $Block['type'] = $blockType;

                    if ( ! isset($Block['identified']))
                    {
                        if (isset($CurrentBlock))
                        {
                            $Elements[] = $this->extractElement($CurrentBlock);
                        }

                        $Block['identified'] = true;
                    }

                    if ($this->isBlockContinuable($blockType))
                    {
                        $Block['continuable'] = true;
                    }

                    $CurrentBlock = $Block;

                    continue 2;
                }
            }

            # ~

            if (isset($CurrentBlock) and $CurrentBlock['type'] === 'Paragraph')
            {
                $Block = $this->paragraphContinue($Line, $CurrentBlock);
            }

            if (isset($Block))
            {
                $CurrentBlock = $Block;
            }
            else
            {
                if (isset($CurrentBlock))
                {
                    $Elements[] = $this->extractElement($CurrentBlock);
                }

                $CurrentBlock = $this->paragraph($Line);

                $CurrentBlock['identified'] = true;
            }
        }

        # ~

        if (isset($CurrentBlock['continuable']) and $this->isBlockCompletable($CurrentBlock['type']))
        {
            $methodName = 'block' . $CurrentBlock['type'] . 'Complete';
            $CurrentBlock = $this->$methodName($CurrentBlock);
        }

        # ~

        if (isset($CurrentBlock))
        {
            $Elements[] = $this->extractElement($CurrentBlock);
        }

        # ~

        return $Elements;
    }

    protected function extractElement(array $Component)
    {
        if ( ! isset($Component['element']))
        {
            if (isset($Component['markup']))
            {
                $Component['element'] = array('rawHtml' => $Component['markup']);
            }
            elseif (isset($Component['hidden']))
            {
                $Component['element'] = array();
            }
        }

        return $Component['element'];
    }

    protected function isBlockContinuable($Type)
    {
        return method_exists($this, 'block' . $Type . 'Continue');
    }

    protected function isBlockCompletable($Type)
    {
        return method_exists($this, 'block' . $Type . 'Complete');
    }

    #
    # Code

    protected function blockCode($Line, $Block = null)
    {
        if (isset($Block) and $Block['type'] === 'Paragraph' and ! isset($Block['interrupted']))
        {
            return;
        }

        if ($Line['indent'] >= 4)
        {
            $text = substr($Line['body'], 4);

            $Block = array(
                'element' => array(
                    'name' => 'pre',
                    'element' => array(
                        'name' => 'code',
                        'text' => $text,
                    ),
                ),
            );

            return $Block;
        }
    }

    protected function blockCodeContinue($Line, $Block)
    {
        if ($Line['indent'] >= 4)
        {
            if (isset($Block['interrupted']))
            {
                $Block['element']['element']['text'] .= str_repeat("\n", $Block['interrupted']);

                unset($Block['interrupted']);
            }

            $Block['element']['element']['text'] .= "\n";

            $text = substr($Line['body'], 4);

            $Block['element']['element']['text'] .= $text;

            return $Block;
        }
    }

    protected function blockCodeComplete($Block)
    {
        return $Block;
    }

    #
    # Comment

    protected function blockComment($Line)
    {
        if ($this->markupEscaped or $this->safeMode)
        {
            return;
        }

        if (strpos($Line['text'], '<!--') === 0)
        {
            $Block = array(
                'element' => array(
                    'comment' => true,
                    'rawHtml' => $Line['body'],
                    'autobreak' => true,
                ),
            );

            if (strpos($Line['text'], '-->') !== false)
            {
                $Block['closed'] = true;
            }

            return $Block;
        }
    }

    protected function blockCommentContinue($Line, array $Block)
    {
        if (isset($Block['closed']))
        {
            return;
        }

        $Block['element']['rawHtml'] .= "\n" . $Line['body'];

        if (strpos($Line['text'], '-->') !== false)
        {
            $Block['closed'] = true;
        }

        return $Block;
    }

    #
    # Fenced Code

    protected function blockFencedCode($Line)
    {
        $marker = $Line['text'][0];

        $openerLength = strspn($Line['text'], $marker);

        if ($openerLength < 3)
        {
            return;
        }

        $infostring = trim(substr($Line['text'], $openerLength), "\t ");

        if (strpos($infostring, '`') !== false)
        {
            return;
        }

        $Element = array(
            'name' => 'code',
            'text' => '',
        );

        if ($infostring !== '')
        {
            /**
             * https://www.w3.org/TR/2011/WD-html5-20110525/elements.html#classes
             * Every HTML element may have a class attribute specified.
             * The attribute, if specified, must have a value that is a set
             * of space-separated tokens representing the various classes
             * that the element belongs to.
             * [...]
             * The space characters, for the purposes of this specification,
             * are U+0020 SPACE, U+0009 CHARACTER TABULATION (tab),
             * U+000A LINE FEED (LF), U+000C FORM FEED (FF), and
             * U+000D CARRIAGE RETURN (CR).
             */
            $language = substr($infostring, 0, strcspn($infostring, " \t\n\f\r"));

            $Element['attributes'] = array('class' => "language-$language");
        }

        $Block = array(
            'char' => $marker,
            'openerLength' => $openerLength,
            'element' => array(
                'name' => 'pre',
                'element' => $Element,
            ),
        );

        return $Block;
    }

    protected function blockFencedCodeContinue($Line, $Block)
    {
        if (isset($Block['complete']))
        {
            return;
        }

        if (isset($Block['interrupted']))
        {
            $Block['element']['element']['text'] .= str_repeat("\n", $Block['interrupted']);

            unset($Block['interrupted']);
        }

        if (($len = strspn($Line['text'], $Block['char'])) >= $Block['openerLength']
            and chop(substr($Line['text'], $len), ' ') === ''
        ) {
            $Block['element']['element']['text'] = substr($Block['element']['element']['text'], 1);

            $Block['complete'] = true;

            return $Block;
        }

        $Block['element']['element']['text'] .= "\n" . $Line['body'];

        return $Block;
    }

    protected function blockFencedCodeComplete($Block)
    {
        return $Block;
    }

    #
    # Header

    protected function blockHeader($Line)
    {
        $level = strspn($Line['text'], '#');

        if ($level > 4)
        {
            return;
        }

        $text = trim($Line['text'], '#');

        if ($this->strictMode and isset($text[0]) and $text[0] !== ' ')
        {
            return;
        }

        $text = trim($text, ' ');

        $Block = array(
            'element' => array(
                'name' => 'h' . $level,
                'text' => $text,
                'handler' => array(
                    'function' => 'lineElements',
                    'argument' => $text,
                    'destination' => 'text',
                )
            ),
        );

        return $Block;
    }

    #
    # List

    protected function blockList($Line, array $CurrentBlock = null)
    {
        list($name, $pattern) = $Line['text'][0] <= '-' ? array('ul', '[*+-]') : array('ol', '[0-9]{1,9}+[.\)]');

        if (preg_match('/^('.$pattern.'([ ]++|$))(.*+)/', $Line['text'], $matches))
        {
            $contentIndent = strlen($matches[2]);

            if ($contentIndent >= 5)
            {
                $contentIndent -= 1;
                $matches[1] = substr($matches[1], 0, -$contentIndent);
                $matches[3] = str_repeat(' ', $contentIndent) . $matches[3];
            }
            elseif ($contentIndent === 0)
            {
                $matches[1] .= ' ';
            }

            $markerWithoutWhitespace = strstr($matches[1], ' ', true);

            $Block = array(
                'indent' => $Line['indent'],
                'pattern' => $pattern,
                'data' => array(
                    'type' => $name,
                    'marker' => $matches[1],
                    'markerType' => ($name === 'ul' ? $markerWithoutWhitespace : substr($markerWithoutWhitespace, -1)),
                ),
                'element' => array(
                    'name' => $name,
                    'elements' => array(),
                ),
            );

            

            $Block['data']['markerTypeRegex'] = preg_quote($Block['data']['markerType'], '/');

            if ($name === 'ol')
            {
                $listStart = ltrim(strstr($matches[1], $Block['data']['markerType'], true), '0') ?: '0';

                if ($listStart !== '1')
                {
                    if (
                        isset($CurrentBlock)
                        and $CurrentBlock['type'] === 'Paragraph'
                        and ! isset($CurrentBlock['interrupted'])
                    ) {
                        return;
                    }

                    $Block['element']['attributes'] = array('start' => $listStart);
                }
            }

            $Block['li'] = array(
                'name' => 'li',
                'handler' => array(
                    'function' => 'li',
                    'argument' => !empty($matches[3]) ? array($matches[3]) : array(),
                    'destination' => 'elements'
                )
            );

            $Block['element']['elements'] []= & $Block['li'];

            return $Block;
        }
    }

    protected function blockListContinue($Line, array $Block)
    {
        if (isset($Block['interrupted']) and empty($Block['li']['handler']['argument']))
        {
            return null;
        }

        $requiredIndent = ($Block['indent'] + strlen($Block['data']['marker']));

        if ($Line['indent'] < $requiredIndent
            and (
                (
                    $Block['data']['type'] === 'ol'
                    and preg_match('/^[0-9]++'.$Block['data']['markerTypeRegex'].'(?:[ ]++(.*)|$)/', $Line['text'], $matches)
                ) or (
                    $Block['data']['type'] === 'ul'
                    and preg_match('/^'.$Block['data']['markerTypeRegex'].'(?:[ ]++(.*)|$)/', $Line['text'], $matches)
                )
            )
        ) {
            if (isset($Block['interrupted']))
            {
                $Block['li']['handler']['argument'] []= '';

                $Block['loose'] = true;

                unset($Block['interrupted']);
            }

            unset($Block['li']);

            $text = isset($matches[1]) ? $matches[1] : '';

            $Block['indent'] = $Line['indent'];

            $Block['li'] = array(
                'name' => 'li',
                'handler' => array(
                    'function' => 'li',
                    'argument' => array($text),
                    'destination' => 'elements'
                )
            );

            $Block['element']['elements'] []= & $Block['li'];

            return $Block;
        }
        elseif ($Line['indent'] < $requiredIndent and $this->blockList($Line))
        {
            return null;
        }

        // Disabled because references aren't needed
        /*
        if ($Line['text'][0] === '[' and $this->blockReference($Line))
        {
            return $Block;
        }
        */

        if ($Line['indent'] >= $requiredIndent)
        {
            if (isset($Block['interrupted']))
            {
                $Block['li']['handler']['argument'] []= '';

                $Block['loose'] = true;

                unset($Block['interrupted']);
            }

            $text = substr($Line['body'], $requiredIndent);

            $Block['li']['handler']['argument'] []= $text;

            return $Block;
        }

        if ( ! isset($Block['interrupted']))
        {
            $text = preg_replace('/^[ ]{0,'.$requiredIndent.'}+/', '', $Line['body']);

            $Block['li']['handler']['argument'] []= $text;

            return $Block;
        }
    }

    protected function blockListComplete(array $Block)
    {
        if (isset($Block['loose']))
        {
            foreach ($Block['element']['elements'] as &$li)
            {
                if (end($li['handler']['argument']) !== '')
                {
                    $li['handler']['argument'] []= '';
                }
            }
        }

        return $Block;
    }

    #
    # Quote

    protected function blockQuote($Line)
    {
        if (preg_match('/^>[ ]?+(.*+)/', $Line['text'], $matches))
        {
            $Block = array(
                'element' => array(
                    'name' => 'blockquote',
                    'handler' => array(
                        'function' => 'linesElements',
                        'argument' => (array) $matches[1],
                        'destination' => 'elements',
                    )
                ),
            );

            return $Block;
        }
    }

    protected function blockQuoteContinue($Line, array $Block)
    {
        if (isset($Block['interrupted']))
        {
            return;
        }

        if ($Line['text'][0] === '>' and preg_match('/^>[ ]?+(.*+)/', $Line['text'], $matches))
        {
            $Block['element']['handler']['argument'] []= $matches[1];

            return $Block;
        }

        if ( ! isset($Block['interrupted']))
        {
            $Block['element']['handler']['argument'] []= $Line['text'];

            return $Block;
        }
    }

    #
    # Rule

    protected function blockRule($Line)
    {
        $marker = $Line['text'][0];

        if (substr_count($Line['text'], $marker) >= 3 and chop($Line['text'], " $marker") === '')
        {
            $Block = array(
                'element' => array(
                    'name' => 'hr',
                ),
            );

            return $Block;
        }
    }

    #
    # Setext

    protected function blockSetextHeader($Line, array $Block = null)
    {
        if ( ! isset($Block) or $Block['type'] !== 'Paragraph' or isset($Block['interrupted']))
        {
            return;
        }

        if ($Line['indent'] < 4 and chop(chop($Line['text'], ' '), $Line['text'][0]) === '')
        {
            $Block['element']['name'] = $Line['text'][0] === '=' ? 'h1' : 'h2';

            return $Block;
        }
    }

    #
    # Markup

    protected function blockMarkup($Line)
    {
        if ($this->markupEscaped or $this->safeMode)
        {
            return;
        }

        if (preg_match('/^<[\/]?+(\w*)(?:[ ]*+'.$this->regexHtmlAttribute.')*+[ ]*+(\/)?>/', $Line['text'], $matches))
        {
            $element = strtolower($matches[1]);

            if (in_array($element, $this->textLevelElements))
            {
                return;
            }

            $Block = array(
                'name' => $matches[1],
                'element' => array(
                    'rawHtml' => $Line['text'],
                    'autobreak' => true,
                ),
            );

            return $Block;
        }
    }

    protected function blockMarkupContinue($Line, array $Block)
    {
        if (isset($Block['closed']) or isset($Block['interrupted']))
        {
            return;
        }

        $Block['element']['rawHtml'] .= "\n" . $Line['body'];

        return $Block;
    }

    #
    # Paragraph style by name
    #

    protected function blockParagraphStyle($Line)
    {
        if (strpos($Line['text'], ']') !== false
            && preg_match('/^\[((?:[a-zA-Z0-9\-_]+)+?)\][ ]?+(.*+)/', $Line['text'], $matches))
        {
            $Block = array(
                'element' => array(
                    'name' => 'blockquote',
                    'attributes' => array('class' => $matches[1]),
                    'handler' => array(
                        'function' => 'linesElements',
                        'argument' => (array) $matches[2],
                        'destination' => 'elements',
                    )
                ),
            );

            return $Block;
        }
    }

    protected function blockParagraphStyleContinue($Line, array $Block)
    {
        if (isset($Block['interrupted']))
        {
            return;
        }

        $Block['element']['handler']['argument'][] = $Line['text'];

        return $Block;
    }

    #
    # Reference

    protected function blockReference($Line)
    {
        if (strpos($Line['text'], ']') !== false
            and preg_match('/^\[(.+?)\]:[ ]*+<?(\S+?)>?(?:[ ]+["\'(](.+)["\')])?[ ]*+$/', $Line['text'], $matches)
        ) {
            $id = strtolower($matches[1]);

            $Data = array(
                'url' => $matches[2],
                'title' => isset($matches[3]) ? $matches[3] : null,
            );

            $this->DefinitionData['Reference'][$id] = $Data;

            $Block = array(
                'element' => array(),
            );

            return $Block;
        }
    }

    #
    # Table

    protected function blockTable($Line, array $Block = null)
    {
        if ( ! isset($Block) or $Block['type'] !== 'Paragraph' or isset($Block['interrupted']))
        {
            return;
        }

        if (
            strpos($Block['element']['handler']['argument'], '|') === false
            and strpos($Line['text'], '|') === false
            and strpos($Line['text'], ':') === false
            or strpos($Block['element']['handler']['argument'], "\n") !== false
        ) {
            return;
        }

        if (chop($Line['text'], ' -:|.0123456789') !== '')
        {
            return;
        }

        $alignments = array();
        $widths = array();
        $areWidthsSpecified = false;

        $divider = $Line['text'];

        $divider = trim($divider);
        $divider = trim($divider, '|');

        $dividerCells = explode('|', $divider);

        foreach ($dividerCells as $dividerCell)
        {
            $dividerCell = trim($dividerCell);

            if ($dividerCell === '')
            {
                return;
            }

            $alignment = null;
            $width = null;

            if ($dividerCell[0] === ':')
            {
                $alignment = 'left';
            }

            if (substr($dividerCell, - 1) === ':')
            {
                $alignment = $alignment === 'left' ? 'center' : 'right';
            }

            $alignments []= $alignment;

            preg_match('~\d+(?:\.\d+)?~', $dividerCell, $matches); // match float
            if (isset($matches[0]))
            {
                $width = $matches[0];
                $areWidthsSpecified = true;
            }

            $widths[] = $width;
        }

        # ~

        $HeaderElements = array();

        $header = $Block['element']['handler']['argument'];

        $header = trim($header);
        
        // replace all vars with placeholders
        $vars = $this->replaceVarsWithPlaceholders($header);
        
        $header = trim($header, '|');
        
        // extract table styles from header line text
        $bgColor = null;
        $tableStyles = array();
        if (preg_match_all($this->StyleOptionsRegexAnywhere, $header, $tblOptionMatches))
        {
            foreach ($tblOptionMatches[1] as $i => $optionListText)
            {
                $options = array_map('trim', explode(',', $optionListText));
                
                foreach ($options as $option)
                {
                    // The style string is guaranteed to contain
                    // the symbol '='because of the matching regex.
                    list($prop, $val) = array_map('trim', explode('=', $option));
                    
                    if (!in_array($prop, $this->availableTableStyles))
                    {
                        throw new Exception("Unsupported table style '$prop'");
                    }
                    
                    // Save bgColor separately from other table styles.
                    // This is because, for some reason, this property won't work 
                    // as a DOCX table style if unit and width are also set.
                    if ($prop === 'bgColor')
                    {
                        $bgColor = $val;
                    }
                    else
                    {
                        $tableStyles[$prop] = $val;
                    }
                }
                $header = str_replace($tblOptionMatches[0][$i], '', $header);
            }
        }

        $headerCells = explode('|', $header);

        if (count($headerCells) !== count($alignments))
        {
            return;
        }

        $headerRowHasText = false;
        foreach ($headerCells as $index => $headerCell)
        {
            $headerCell = trim($headerCell);
            if (!$headerRowHasText && $headerCell != '')
            {
                $headerRowHasText = true;
            }
            
            // replace placeholders back to variables
            $this->replacePlaceholdersWithVars($headerCell, $vars);

            $HeaderElement = array(
                'name' => 'th',
                'column_alignment' => $alignments[$index],
                'column_width' => $widths[$index],
                'table_bg_color' => $bgColor,
                'table_styles' => $tableStyles,
                'handler' => array(
                    'function' => 'lineElements',
                    'argument' => $headerCell,
                    'destination' => 'elements',
                )
            );

            if (isset($alignments[$index]) || isset($widths[$index]))
            {
                $HeaderElement['attributes']['style'] = '';
                if (isset($alignments[$index]))
                {
                    $alignment = $alignments[$index];
                    $HeaderElement['attributes']['style'] .= "text-align: $alignment;";
                }
    
                if (isset($widths[$index]))
                {
                    $width = $widths[$index];
                    $HeaderElement['attributes']['style'] .= 'width: '.\PhpOffice\PhpWord\Shared\Converter::cmToPixel($width).'px;';
                }
            }

            $HeaderElements []= $HeaderElement;
        }

        # ~

        $Block = array(
            'column_alignments' => $alignments,
            'column_widths' => $widths,
            'table_bg_color' => $bgColor,
            'table_styles' => $tableStyles,
            'identified' => true,
            'element' => array(
                'name' => 'table',
                'are_widths_specified' => $areWidthsSpecified,
                'column_widths' => $widths,
                'bg_color' => $bgColor,
                'row_num' => $headerRowHasText ? 1 : 0,
                'column_num' => count($headerCells),
                'elements' => array(),
                'table_styles' => $tableStyles,
            ),
        );

        $Block['element']['elements'] []= array(
            'name' => 'thead',
        );

        $Block['element']['elements'] []= array(
            'name' => 'tbody',
            'elements' => array(),
        );

        $Block['element']['elements'][0]['elements'] []= array(
            'name' => 'tr',
            'elements' => $HeaderElements,
        );

        return $Block;
    }

    protected function blockTableContinue($Line, array $Block)
    {
        if (isset($Block['interrupted']))
        {
            return;
        }

        if (count($Block['column_alignments']) === 1 or $Line['text'][0] === '|' or strpos($Line['text'], '|'))
        {
            $Elements = array();

            $row = $Line['text'];

            $row = trim($row);
            
            // replace all vars with placeholders
            $vars = $this->replaceVarsWithPlaceholders($row);
            
            $row = trim($row, '|');

            preg_match_all('/(?:(\\\\[|])|[^|`]|`[^`]++`|`)++/', $row, $matches);

            $cells = array_slice($matches[0], 0, count($Block['column_alignments']));

            foreach ($cells as $index => $cell)
            {
                $cell = trim($cell);
                
                // extract cell styles from cell text
                $cellStyles = array();
                if (preg_match_all($this->StyleOptionsRegexAnywhere, $cell, $cellStyleMatches))
                {
                    foreach ($cellStyleMatches[1] as $i => $styleListText)
                    {
                        $styles = array_map('trim', explode(',', $styleListText));
                        
                        foreach ($styles as $style)
                        {
                            // The style string is guaranteed to contain
                            // the symbol '='because of the matching regex.
                            list($prop, $val) = array_map('trim', explode('=', $style));
                            
                            if (!in_array($prop, $this->availableCellStyles))
                            {
                                throw new Exception("Unsupported table cell style '$prop'");
                            }
                            
                            $cellStyles[$prop] = $val;
                        }
                        $cell = str_replace($cellStyleMatches[0][$i], '', $cell);
                    }
                }
                
                // replace placeholders back to variables
                $this->replacePlaceholdersWithVars($cell, $vars);

                $Element = array(
                    'name' => 'td',
                    'column_alignment' => $Block['column_alignments'][$index],
                    'column_width' => $Block['column_widths'][$index],
                    'table_bg_color' => $Block['table_bg_color'],
                    'table_styles' => $Block['table_styles'],
                    'cell_styles' => $cellStyles,
                    'handler' => array(
                        'function' => 'lineElements',
                        'argument' => $cell,
                        'destination' => 'elements',
                    )
                );

                if (isset($Block['column_alignments'][$index]) || isset($Block['column_widths'][$index]))
                {
                    $Element['attributes']['style'] = '';
                    if (isset($Block['column_alignments'][$index]))
                    {
                        $Element['attributes']['style'] .= 'text-align: ' . $Block['column_alignments'][$index] . ';';
                    }
    
                    if (isset($Block['column_widths'][$index]))
                    {
                        $Element['attributes']['style'] .= 'width: ' . \PhpOffice\PhpWord\Shared\Converter::cmToPixel($Block['column_widths'][$index]) . 'px;';
                    }
                }

                $Elements []= $Element;
            }

            $Element = array(
                'name' => 'tr',
                'elements' => $Elements,
            );

            $Block['element']['row_num'] += 1;

            $Block['element']['elements'][1]['elements'] []= $Element;

            return $Block;
        }
    }

    #
    # ~
    #

    protected function paragraph($Line)
    {
        return array(
            'type' => 'Paragraph',
            'element' => array(
                'name' => 'p',
                'handler' => array(
                    'function' => 'lineElements',
                    'argument' => $Line['text'],
                    'destination' => 'elements',
                ),
            ),
        );
    }

    protected function paragraphContinue($Line, array $Block)
    {
        if (isset($Block['interrupted']))
        {
            return;
        }

        $Block['element']['handler']['argument'] .= "\n".$Line['text'];

        return $Block;
    }

    #
    # Inline Elements
    #

    protected $InlineTypes = array(
        '!' => array('Image'),
        '&' => array('SpecialCharacter'),
        '*' => array('Emphasis'),
        ':' => array('Url'),
        '<' => array('UrlTag', 'EmailTag', 'Markup'),
        '[' => array('Link'),
        '_' => array('Emphasis'),
        // '`' => array('Code'),
        '`' => array('FontStyle'),
        '~' => array('Strikethrough'),
        '\\' => array('EscapeSequence'),
        '$' => array('Var'),
    );

    # ~

    protected $inlineMarkerList = '!*_&[:<`~\\$';

    #
    # Process single-line text
    #

    public function line($text, $nonNestables = array())
    {
        $lineElements = $this->lineElements($text, $nonNestables);

        if (!$this->outputMode)
        {
            throw new Exception('Output mode not set');
        }

        if ($this->outputMode == self::TYPE_ODT)
        {
            $this->elementsOdt($lineElements);
        }
        elseif ($this->outputMode == self::TYPE_DOCX)
        {
            if ($this->section == null)
            {
                $this->section = $this->phpWord->addSection();
                $this->sectionIdx += 1;
                $this->phpWord->addNumberingStyle('roman'.$this->sectionIdx, $this->romanNumberingStyle);
                $this->phpWord->addNumberingStyle('decimal'.$this->sectionIdx, $this->decimalNumberingStyle);
                $this->setPhpWordMargins();
            }
            $this->elementsDocx($lineElements);
        }
        elseif ($this->outputMode == self::TYPE_HTML)
        {
            $this->elements($lineElements);
        }
    }
    
    protected function lineVars($text)
    {
        # standardize line breaks
        $text = str_replace(array("\r\n", "\r"), "\n", $text);
        
        $Elements = array();
        
        # $excerpt is based on the first occurrence of a marker
        
        while ($excerpt = strpbrk($text, "$"))
        {
            $marker = $excerpt[0];
            
            $markerPosition = strlen($text) - strlen($excerpt);
            
            $Excerpt = array('text' => $excerpt, 'context' => $text);
            
            foreach ($this->InlineTypes[$marker] as $inlineType)
            {                
                $Inline = $this->{"inline$inlineType"}($Excerpt);
                
                if ( ! isset($Inline))
                {
                    continue;
                }
                
                # makes sure that the inline belongs to "our" marker
                
                if (isset($Inline['position']) and $Inline['position'] > $markerPosition)
                {
                    continue;
                }
                
                # sets a default inline position
                
                if ( ! isset($Inline['position']))
                {
                    $Inline['position'] = $markerPosition;
                }
                
                # the text that comes before the inline
                $unmarkedText = substr($text, 0, $Inline['position']);
                
                # compile the unmarked text
                $InlineText = $this->inlineText($unmarkedText);
                $Elements[] = $InlineText['element'];
                
                # compile the inline
                $Elements[] = $this->extractElement($Inline);
                
                # remove the examined text
                $text = substr($text, $Inline['position'] + $Inline['extent']);
                
                continue 2;
            }
            
            # the marker does not belong to an inline
            
            $unmarkedText = substr($text, 0, $markerPosition + 1);
            
            $InlineText = $this->inlineText($unmarkedText);
            $Elements[] = $InlineText['element'];
            
            $text = substr($text, $markerPosition + 1);
        }
        
        $InlineText = $this->inlineText($text);
        $Elements[] = $InlineText['element'];
        
        foreach ($Elements as &$Element)
        {
            if ( ! isset($Element['autobreak']))
            {
                $Element['autobreak'] = false;
            }
        }
        
        return $Elements;
    }

    protected function lineElements($text, $nonNestables = array())
    {
        # standardize line breaks
        $text = str_replace(array("\r\n", "\r"), "\n", $text);

        $Elements = array();

        $nonNestables = (empty($nonNestables)
            ? array()
            : array_combine($nonNestables, $nonNestables)
        );

        # $excerpt is based on the first occurrence of a marker

        while ($excerpt = strpbrk($text, $this->inlineMarkerList))
        {
            $marker = $excerpt[0];

            $markerPosition = strlen($text) - strlen($excerpt);

            $Excerpt = array('text' => $excerpt, 'context' => $text);

            foreach ($this->InlineTypes[$marker] as $inlineType)
            {
                # check to see if the current inline type is nestable in the current context

                if (isset($nonNestables[$inlineType]))
                {
                    continue;
                }

                $Inline = $this->{"inline$inlineType"}($Excerpt);

                if ( ! isset($Inline))
                {
                    continue;
                }

                # makes sure that the inline belongs to "our" marker

                if (isset($Inline['position']) and $Inline['position'] > $markerPosition)
                {
                    continue;
                }

                # sets a default inline position

                if ( ! isset($Inline['position']))
                {
                    $Inline['position'] = $markerPosition;
                }

                # cause the new element to 'inherit' our non nestables


                $Inline['element']['nonNestables'] = isset($Inline['element']['nonNestables'])
                    ? array_merge($Inline['element']['nonNestables'], $nonNestables)
                    : $nonNestables
                ;

                # the text that comes before the inline
                $unmarkedText = substr($text, 0, $Inline['position']);

                # compile the unmarked text
                $InlineText = $this->inlineText($unmarkedText);
                $Elements[] = $InlineText['element'];

                # compile the inline
                $Elements[] = $this->extractElement($Inline);

                # remove the examined text
                $text = substr($text, $Inline['position'] + $Inline['extent']);

                continue 2;
            }

            # the marker does not belong to an inline

            $unmarkedText = substr($text, 0, $markerPosition + 1);

            $InlineText = $this->inlineText($unmarkedText);
            $Elements[] = $InlineText['element'];

            $text = substr($text, $markerPosition + 1);
        }

        $InlineText = $this->inlineText($text);
        $Elements[] = $InlineText['element'];

        foreach ($Elements as &$Element)
        {
            if ( ! isset($Element['autobreak']))
            {
                $Element['autobreak'] = false;
            }
        }

        return $Elements;
    }

    #
    # ~
    #

    protected function inlineText($text)
    {
        $Inline = array(
            'extent' => strlen($text),
            'element' => array(),
        );

        $Inline['element']['elements'] = self::pregReplaceElements(
            $this->breaksEnabled ? '/[ ]*+\n/' : '/(?:[ ]*+\\\\|[ ]{2,}+)\n/',
            array(
                array('name' => 'br'),
                array('text' => "\n"),
            ),
            $text
        );

        return $Inline;
    }

    protected function inlineFontStyle($Excerpt)
    {
        $marker = $Excerpt['text'][0];

        if ( ! isset($Excerpt['text'][1]) or $Excerpt['text'][1] !== '[')
        {
            return;
        }

        if (!preg_match($this->fontStyleRegex, $Excerpt['text'], $matches))
        {
            return;
        }
        
        if (!$matches[1] && !$matches[2] && !$matches[3])
        {
            return;
        }

        return array(
            'extent' => strlen($matches[0]),
            'element' => array(
                'name' => 'fontstyle',
                'color' => $matches[1],
                'bgColor' => $matches[2],
                'size' => $matches[3],
                'handler' => array(
                    'function' => 'lineElements',
                    'argument' => $matches[4],
                    'destination' => 'elements',
                )
            ),
        );
    }

    protected function inlineCode($Excerpt)
    {
        $marker = $Excerpt['text'][0];

        if (preg_match('/^(['.$marker.']++)[ ]*+(.+?)[ ]*+(?<!['.$marker.'])\1(?!'.$marker.')/s', $Excerpt['text'], $matches))
        {
            $text = $matches[2];
            $text = preg_replace('/[ ]*+\n/', ' ', $text);

            return array(
                'extent' => strlen($matches[0]),
                'element' => array(
                    'name' => 'code',
                    'text' => $text,
                ),
            );
        }
    }

    protected function inlineEmailTag($Excerpt)
    {
        $hostnameLabel = '[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?';

        $commonMarkEmail = '[a-zA-Z0-9.!#$%&\'*+\/=?^_`{|}~-]++@'
            . $hostnameLabel . '(?:\.' . $hostnameLabel . ')*';

        if (strpos($Excerpt['text'], '>') !== false
            and preg_match("/^<((mailto:)?$commonMarkEmail)>/i", $Excerpt['text'], $matches)
        ){
            $url = $matches[1];

            if ( ! isset($matches[2]))
            {
                $url = "mailto:$url";
            }

            return array(
                'extent' => strlen($matches[0]),
                'element' => array(
                    'name' => 'a',
                    'text' => $matches[1],
                    'attributes' => array(
                        'href' => $url,
                    ),
                ),
            );
        }
    }

    protected function inlineEmphasis($Excerpt)
    {
        if ( ! isset($Excerpt['text'][1]))
        {
            return;
        }

        $marker = $Excerpt['text'][0];
        
        // replace all vars with placeholders
        $vars = $this->replaceVarsWithPlaceholders($Excerpt['text']);

        if ($Excerpt['text'][1] === $marker and preg_match($this->StrongRegex['*'], $Excerpt['text'], $matches))
        {
            $emphasis = 'strong';
        }
        elseif (preg_match($this->EmRegex['*'], $Excerpt['text'], $matches))
        {
            $emphasis = 'em';
        }
        elseif (preg_match($this->UnderlineRegex, $Excerpt['text'], $matches))
        {
            $emphasis = 'u';
        }
        else
        {
            return;
        }
        
        // replace placeholders back to variables
        foreach ($matches as &$match)
        {
            $this->replacePlaceholdersWithVars($match, $vars);
        }

        return array(
            'extent' => strlen($matches[0]),
            'element' => array(
                'name' => $emphasis,
                'handler' => array(
                    'function' => 'lineElements',
                    'argument' => $matches[1],
                    'destination' => 'elements',
                )
            ),
        );
    }

    protected function inlineVar($Excerpt)
    {
        if ( ! isset($Excerpt['text'][1]))
        {
            return;
        }

        $marker = $Excerpt['text'][0];

        if ($Excerpt['text'][1] === '{' and preg_match($this->VariableRegex, $Excerpt['text'], $matches))
        {
            $name = 'var';
        }
        else
        {
            return;
        }

        $varPattern = $matches[1];
        $parts = explode('|', $varPattern);
        $key = array_shift($parts);

        $params = array();
        foreach ($parts as $part)
        {
            $params[] = explode(',', $part);
        }

        return array(
            'extent' => strlen($matches[0]),
            'element' => array(
                'name' => $name,
                'key' => $key,
                'params' => $params,
                'text' => '${'.$varPattern.'}',
            ),
        );
    }

    protected function inlineEscapeSequence($Excerpt)
    {
        if (isset($Excerpt['text'][1]) and in_array($Excerpt['text'][1], $this->specialCharacters))
        {
            return array(
                'element' => array('rawHtml' => $Excerpt['text'][1]),
                'extent' => 2,
            );
        }
    }

    protected function inlineImage($Excerpt)
    {
        if ( ! isset($Excerpt['text'][1]) or $Excerpt['text'][1] !== '[')
        {
            return;
        }

        $Excerpt['text']= substr($Excerpt['text'], 1);

        $Link = $this->inlineLink($Excerpt);

        if ($Link === null)
        {
            return;
        }

        $Inline = array(
            'extent' => $Link['extent'] + 1,
            'element' => array(
                'name' => 'img',
                'attributes' => array(
                    'src' => $Link['element']['attributes']['href'],
                    'alt' => $Link['element']['handler']['argument'],
                ),
                'autobreak' => true,
            ),
        );

        $Inline['element']['attributes'] += $Link['element']['attributes'];

        unset($Inline['element']['attributes']['href']);

        return $Inline;
    }

    protected function inlineLink($Excerpt)
    {
        $Element = array(
            'name' => 'a',
            'handler' => array(
                'function' => 'lineElements',
                'argument' => null,
                'destination' => 'elements',
            ),
            'nonNestables' => array('Url', 'Link'),
            'attributes' => array(
                'href' => null,
                'title' => null,
            ),
        );

        $extent = 0;

        $remainder = $Excerpt['text'];

        if (preg_match('/\[((?:[^][]++|(?R))*+)\]/', $remainder, $matches))
        {
            $Element['handler']['argument'] = $matches[1];

            $extent += strlen($matches[0]);

            $remainder = substr($remainder, $extent);
        }
        else
        {
            return;
        }

        if (preg_match('/^[(]\s*+((?:[^ ()]++|[(][^ )]+[)])++)(?:[ ]+("[^"]*+"|\'[^\']*+\'))?\s*+[)]/', $remainder, $matches))
        {
            $Element['attributes']['href'] = $matches[1];

            if (isset($matches[2]))
            {
                $Element['attributes']['title'] = substr($matches[2], 1, - 1);
            }

            $extent += strlen($matches[0]);
        }
        else
        {
            if (preg_match('/^\s*\[(.*?)\]/', $remainder, $matches))
            {
                $definition = strlen($matches[1]) ? $matches[1] : $Element['handler']['argument'];
                $definition = strtolower($definition);

                $extent += strlen($matches[0]);
            }
            else
            {
                $definition = strtolower($Element['handler']['argument']);
            }

            if ( ! isset($this->DefinitionData['Reference'][$definition]))
            {
                return;
            }

            $Definition = $this->DefinitionData['Reference'][$definition];

            $Element['attributes']['href'] = $Definition['url'];
            $Element['attributes']['title'] = $Definition['title'];
        }

        return array(
            'extent' => $extent,
            'element' => $Element,
        );
    }

    protected function inlineMarkup($Excerpt)
    {
        if ($this->markupEscaped or $this->safeMode or strpos($Excerpt['text'], '>') === false)
        {
            return;
        }

        if ($Excerpt['text'][1] === '/' and preg_match('/^<\/\w[\w-]*+[ ]*+>/s', $Excerpt['text'], $matches))
        {
            return array(
                'element' => array('rawHtml' => $matches[0]),
                'extent' => strlen($matches[0]),
            );
        }

        if ($Excerpt['text'][1] === '!' and preg_match('/^<!---?[^>-](?:-?+[^-])*-->/s', $Excerpt['text'], $matches))
        {
            return array(
                'element' => array('rawHtml' => $matches[0]),
                'extent' => strlen($matches[0]),
            );
        }

        if ($Excerpt['text'][1] !== ' ' and preg_match('/^<\w[\w-]*+(?:[ ]*+'.$this->regexHtmlAttribute.')*+[ ]*+\/?>/s', $Excerpt['text'], $matches))
        {
            return array(
                'element' => array('rawHtml' => $matches[0]),
                'extent' => strlen($matches[0]),
            );
        }
    }

    protected function inlineSpecialCharacter($Excerpt)
    {
        if (substr($Excerpt['text'], 1, 1) !== ' ' and strpos($Excerpt['text'], ';') !== false
            and preg_match('/^&(#?+[0-9a-zA-Z]++);/', $Excerpt['text'], $matches)
        ) {
            return array(
                'element' => array('rawHtml' => '&' . $matches[1] . ';'),
                'extent' => strlen($matches[0]),
            );
        }

        return;
    }

    protected function inlineStrikethrough($Excerpt)
    {
        if ( ! isset($Excerpt['text'][1]))
        {
            return;
        }

        if ($Excerpt['text'][1] === '~' and preg_match('/^~~(?=\S)(.+?)(?<=\S)~~/', $Excerpt['text'], $matches))
        {
            return array(
                'extent' => strlen($matches[0]),
                'element' => array(
                    'name' => 'del',
                    'handler' => array(
                        'function' => 'lineElements',
                        'argument' => $matches[1],
                        'destination' => 'elements',
                    )
                ),
            );
        }
    }

    protected function inlineUrl($Excerpt)
    {
        if ($this->urlsLinked !== true or ! isset($Excerpt['text'][2]) or $Excerpt['text'][2] !== '/')
        {
            return;
        }

        if (strpos($Excerpt['context'], 'http') !== false
            and preg_match('/\bhttps?+:[\/]{2}[^\s<]+\b\/*+/ui', $Excerpt['context'], $matches, PREG_OFFSET_CAPTURE)
        ) {
            $url = $matches[0][0];

            $Inline = array(
                'extent' => strlen($matches[0][0]),
                'position' => $matches[0][1],
                'element' => array(
                    'name' => 'a',
                    'text' => $url,
                    'attributes' => array(
                        'href' => $url,
                    ),
                ),
            );

            return $Inline;
        }
    }

    protected function inlineUrlTag($Excerpt)
    {
        if (strpos($Excerpt['text'], '>') !== false and preg_match('/^<(\w++:\/{2}[^ >]++)>/i', $Excerpt['text'], $matches))
        {
            $url = $matches[1];

            return array(
                'extent' => strlen($matches[0]),
                'element' => array(
                    'name' => 'a',
                    'text' => $url,
                    'attributes' => array(
                        'href' => $url,
                    ),
                ),
            );
        }
    }

    # ~

    protected function unmarkedText($text)
    {
        $Inline = $this->inlineText($text);
        return $this->element($Inline['element']);
    }

    #
    # Handlers
    #

    protected function handle(array $Element)
    {
        if (isset($Element['handler']))
        {
            if (!isset($Element['nonNestables']))
            {
                $Element['nonNestables'] = array();
            }

            if (is_string($Element['handler']))
            {
                $function = $Element['handler'];
                $argument = $Element['text'];
                unset($Element['text']);
                $destination = 'rawHtml';
            }
            else
            {
                $function = $Element['handler']['function'];
                $argument = $Element['handler']['argument'];
                $destination = $Element['handler']['destination'];
            }

            if ($destination === 'text') {
                // handles variables in headings
                $varElements = $this->{$function}($argument);
                foreach ($varElements as $element)
                {
                    if (isset($element['name']) && $element['name'] === 'var')
                    {
                        $value = $this->varConverter->evaluate($element['key'], $element['params']);
                        
                        if (isset($value['text'])) {
                            $value = $value['text'];
                        }
                        
                        if ($this->canBeString($value)) {
                            $Element[$destination] = str_replace($element['text'], $value, $Element[$destination]);
                        }
                    }
                }
            }
            else
            {
                $Element[$destination] = $this->{$function}($argument, $Element['nonNestables']);
                
                if ($destination === 'handler')
                {
                    $Element = $this->handle($Element);
                }
            }
            
            unset($Element['handler']);
        }

        return $Element;
    }

    protected function handleElementRecursive(array $Element)
    {
        return $this->elementApplyRecursive(array($this, 'handle'), $Element);
    }

    protected function handleElementsRecursive(array $Elements)
    {
        return $this->elementsApplyRecursive(array($this, 'handle'), $Elements);
    }

    protected function elementApplyRecursive($closure, array $Element)
    {
        $Element = call_user_func($closure, $Element);

        if (isset($Element['elements']))
        {
            $Element['elements'] = $this->elementsApplyRecursive($closure, $Element['elements']);
        }
        elseif (isset($Element['element']))
        {
            $Element['element'] = $this->elementApplyRecursive($closure, $Element['element']);
        }

        return $Element;
    }

    protected function elementApplyRecursiveDepthFirst($closure, array $Element)
    {
        if (isset($Element['elements']))
        {
            $Element['elements'] = $this->elementsApplyRecursiveDepthFirst($closure, $Element['elements']);
        }
        elseif (isset($Element['element']))
        {
            $Element['element'] = $this->elementsApplyRecursiveDepthFirst($closure, $Element['element']);
        }

        $Element = call_user_func($closure, $Element);

        return $Element;
    }

    protected function elementsApplyRecursive($closure, array $Elements)
    {
        foreach ($Elements as &$Element)
        {
            $Element = $this->elementApplyRecursive($closure, $Element);
        }

        return $Elements;
    }

    protected function elementsApplyRecursiveDepthFirst($closure, array $Elements)
    {
        foreach ($Elements as &$Element)
        {
            $Element = $this->elementApplyRecursiveDepthFirst($closure, $Element);
        }

        return $Elements;
    }

    protected function element(array $Element)
    {
        if ($this->safeMode)
        {
            $Element = $this->sanitiseElement($Element);
        }

        # identity map if element has no handler
        $Element = $this->handle($Element);

        $hasName = isset($Element['name']);

        if ($hasName)
        {
            $this->markup .= '<' . $Element['name'];

            if (isset($Element['attributes']))
            {
                foreach ($Element['attributes'] as $name => $value)
                {
                    if ($value === null)
                    {
                        continue;
                    }

                    $this->markup .= " $name=\"".self::escape($value).'"';
                }
            }
        }

        $permitRawHtml = false;

        if (isset($Element['text']))
        {
            $text = $Element['text'];
        }
        // very strongly consider an alternative if you're writing an
        // extension
        elseif (isset($Element['rawHtml']))
        {
            $text = $Element['rawHtml'];

            $allowRawHtmlInSafeMode = isset($Element['allowRawHtmlInSafeMode']) && $Element['allowRawHtmlInSafeMode'];
            $permitRawHtml = !$this->safeMode || $allowRawHtmlInSafeMode;
        }

        $hasContent = isset($text) || isset($Element['element']) || isset($Element['elements']);

        if ($hasContent)
        {
            $this->markup .= $hasName ? '>' : '';

            if (isset($Element['elements']))
            {
                $this->elements($Element['elements']);
            }
            elseif (isset($Element['element']))
            {
                $this->element($Element['element']);
            }
            else
            {
                if ($hasName && $Element['name'] == 'var')
                {
                    $value = $this->varConverter->evaluate($Element['key'], $Element['params']);

                    // handle variable value 
                    if (is_array($value))
                    {
                        // value is array
                        if (isset($value['markdown'], $value['text']) && $value['text'])
                        {
                            if ($value['markdown'])
                            {
                                // value is markdown text, so parse it and handle elements
                                $tree = $this->textElements($value['text']);
                                if (isset($tree[0]['handler']['argument']))
                                {
                                    $md = $tree[0]['handler']['argument'];
                                    if (strstr($md, PHP_EOL))
                                    {
                                        // multi-line string
                                        $this->text($md);
                                    }
                                    else
                                    {
                                        // single-line string
                                        $this->line($md);
                                    }
                                }
                                else
                                {
                                    $this->elements($tree);
                                }
                            }
                            else
                            {
                                $this->markup .= $value['text'];
                            }
                        }
                    }
                    else
                    {
                        $this->markup .= $value;
                    }
                }
                else
                {
                    if (!$permitRawHtml)
                    {
                        $this->markup .= self::escape($text, true);
                    }
                    else
                    {
                        $this->markup .= $text;
                    }
                }
            }

            $this->markup .= $hasName ? '</' . $Element['name'] . '>' : '';
        }
        elseif ($hasName)
        {
            $this->markup .= ' />';
        }
    }

    protected function elementDocx(array $Element)
    {
        /*
        if ($this->safeMode)
        {
            $Element = $this->sanitiseElement($Element);
        }
        */

        # identity map if element has no handler
        $Element = $this->handle($Element);

        $hasName = isset($Element['name']);

        if ($hasName)
        {
            if ($Element['name'] == 'p')
            {
                // do nothing!
                // text runs are created on the fly
            }
            elseif ($Element['name'] == 'strong')
            {
                $this->options['bold'] = true;
            }
            elseif ($Element['name'] == 'em')
            {
                $this->options['italic'] = true;
            }
            elseif ($Element['name'] == 'u')
            {
                $this->options['underline'] = 'single';
            }
            elseif ($Element['name'] == 'img')
            {
                if (isset($Element['attributes']['src']))
                {
                    $options = array();
                    if (isset($Element['attributes']['alt']) && $Element['attributes']['alt'])
                    {
                        $altArgs = explode(',', $Element['attributes']['alt']);
                        $countArgs = count($altArgs);
                        $height = 0;
                        $width = 0;
                        if ($countArgs >= 2)
                        {
                            $height = $altArgs[0];
                            $width = $altArgs[1];
                        }
                        elseif ($countArgs == 1)
                        {
                            $height = $altArgs[0];
                        }
                        $options['height'] = \PhpOffice\PhpWord\Shared\Converter::cmToPoint($height);
                        if ($width)
                        {
                            $options['width'] = \PhpOffice\PhpWord\Shared\Converter::cmToPoint($width);
                        }
                    }
                    
                    $run = $this->getRun();
                    $run->addImage($Element['attributes']['src'], $options);
                }
            }
            elseif (in_array($Element['name'], array('h1', 'h2', 'h3', 'h4')))
            {
                switch ($Element['name'])
                {
                    case 'h2': $titleLevel = $this->headingsMap[2]; break;
                    case 'h3': $titleLevel = $this->headingsMap[3]; break;
                    case 'h4': $titleLevel = $this->headingsMap[4]; break;
                    default: $titleLevel = $this->headingsMap[1];
                }
                $this->section->addTitle($Element['text'], $titleLevel);
            }
            elseif (in_array($Element['name'], array('ol', 'ul')))
            {
                $parentListStyle = $this->listStyleName;
                $this->listDepth += 1;
                $this->listStyleName = $Element['name'] == 'ol' ? ($this->listIsRoman ? 'roman'.$this->sectionIdx : 'decimal'.$this->sectionIdx) : ($this->listIsBullets ? 'bullet' : 'hyphen');
            }
            elseif ($Element['name'] == 'li')
            {
                $this->listItemRun = $this->section->addListItemRun($this->listDepth, $this->listStyleName);
            }
            elseif ($Element['name'] == 'hr')
            {
                $this->section->addTextBreak();
            }
            elseif ($Element['name'] == 'br')
            {
                if ($this->listItemRun != null)
                {
                    $this->listItemRun->addTextBreak();
                }
                elseif ($this->textRun != null)
                {
                    $this->textRun->addTextBreak();
                }
                else
                {
                    $this->section->addTextBreak();
                }
            }
            elseif ($Element['name'] == 'blockquote')
            {
                if (isset($Element['attributes']['class']))
                {
                    $this->pStyle = $Element['attributes']['class'];
                }
                else
                {
                    $this->paragraphDepth += 1;
                }
            }
            elseif ($Element['name'] == 'table')
            {
                $tableStyle = $Element['table_styles'];
                if (!$Element['are_widths_specified'])
                {
                    // if column widths not specified, set the table width to 100%
                    $tableStyle['unit'] = \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT;
                    $tableStyle['width'] = 50 * 100;
                }
                
                $cellMarginProps = array('cellMargin'/*, 'cellMarginTop', 'cellMarginBottom', 'cellMarginRight', 'cellMarginLeft'*/);
                foreach ($cellMarginProps as $prop)
                {
                    if (isset($tableStyle[$prop]))
                    {
                        // convert cm to twip
                        $tableStyle[$prop] = \PhpOffice\PhpWord\Shared\Converter::cmToTwip($tableStyle[$prop]);
                    }
                }
                
                $this->table = $this->section->addTable($tableStyle);
            }
            elseif ($Element['name'] == 'tr')
            {
                // Check if text of table headers exists.
                // If no text supplied, the table header row won't be created.
                $thHasText = false;
                if ($this->isTh && $Element['elements'])
                {
                    foreach ($Element['elements'] as $el)
                    {
                        if ($el['name'] == 'th' && trim($el['handler']['argument']) !== '')
                        {
                            $thHasText = true;
                            break;
                        }
                    }
                }

                if (!$this->isTh || $thHasText)
                {
                    $this->row = $this->table->addRow();
                }
                
            }
            elseif (in_array($Element['name'], array('td', 'th')))
            {
                // Create cell only if row exists.
                // Row won't exist if no table header text supplied.
                // Also skip cell creation if colSpan is in effect.
                if ($this->row)
                {
                    $this->addTextToCell = false;
                    
                    if ($this->currentColSpan == 0)
                    {
                        $cellStyle = array();
                        if ($Element['name'] === 'td')
                        {
                            $cellStyle = $Element['cell_styles'];
                        }
                        
                        // set cell alignment
                        $alignment = 'left';
                        if (isset($Element['column_alignment']) || isset($cellStyle['textAlignment']))
                        {
                            if (isset($cellStyle['textAlignment']))
                            {
                                $alignment = $cellStyle['textAlignment'];
                            }
                            else
                            {
                                // column alignment as fallback
                                $alignment = $Element['column_alignment'];
                            }
                        }
                        switch ($alignment)
                        {
                            case 'left': $this->pStyleCell = 'left_aligned'; break;
                            case 'right': $this->pStyleCell = 'right_aligned'; break;
                            case 'center': $this->pStyleCell = 'center_aligned'; break;
                        }
                        
                        // Set cell width.
                        // Original solution:
                        // In cases when column width is not specified,
                        // set cell width to null (default value) when creating the table cell,
                        // so the width is determined automatically.
                        // DOCX files generated this way opened correctly in MS Word,
                        // but LibreOffice reported an error when trying to open them.
                        // $cellWidth = isset($Element['column_width']) ? \PhpOffice\PhpWord\Shared\Converter::cmToTwip($Element['column_width']) : null;

                        // Alternative solution (workaround):
                        // In those cases, instead of setting the cell width to null,
                        // set it to any positive number.
                        $cellWidth = isset($Element['column_width']) ? $Element['column_width'] : 5;
                        $cellWidth = \PhpOffice\PhpWord\Shared\Converter::cmToTwip($cellWidth);
                        
                        // set cell background color to table background color as fallback if exists
                        if (!isset($cellStyle['bgColor']) && isset($Element['table_bg_color']))
                        {
                            $cellStyle['bgColor'] = $Element['table_bg_color'];
                        }
                        
                        if (isset($cellStyle['gridSpan']))
                        {
                            $this->currentColSpan = $cellStyle['gridSpan'];
                        }
                        
                        $this->cell = $this->row->addCell($cellWidth, $cellStyle);
                        $this->textRun = $this->cell->addTextRun($this->pStyleCell);
                        $this->addTextToCell = true;
                    }
                }
            }
            elseif ($Element['name'] == 'thead')
            {
                $this->isTh = true;
            }
            elseif ($Element['name'] == 'tbody')
            {
                $this->isTh = false;
            }
            elseif ($Element['name'] == 'fontstyle')
            {
                $this->options['color'] = $Element['color'];
                $this->options['bgColor'] = $Element['bgColor'];
                $this->options['size'] = $Element['size'];
            }
        }

        $permitRawHtml = false;

        if (isset($Element['text']))
        {
            $text = $Element['text'];
        }

        
        // very strongly consider an alternative if you're writing an
        // extension
        elseif (isset($Element['rawHtml']))
        {
            $text = $Element['rawHtml'];

            $allowRawHtmlInSafeMode = isset($Element['allowRawHtmlInSafeMode']) && $Element['allowRawHtmlInSafeMode'];
            $permitRawHtml = !$this->safeMode || $allowRawHtmlInSafeMode;
        }
        

        $hasContent = isset($text) || isset($Element['element']) || isset($Element['elements']);

        if ($hasContent)
        {
            if (isset($Element['elements']))
            {
                $this->elementsDocx($Element['elements']);
            }
            elseif (isset($Element['element']))
            {
                $this->elementDocx($Element['element']);
            }
            else
            {
                if ($hasName)
                {
                    if ($Element['name'] == 'var')
                    {
                        // handle variable value
                        
                        $value = $this->varConverter->evaluate($Element['key'], $Element['params']);

                        if (is_array($value) && isset($value['markdown'], $value['text']))
                        {
                            // value is array
                            
                            if ($value['text'] != '')
                            {
                                if ($value['markdown'])
                                {
                                    // value is markdown text, so parse it and handle elements
                                    if (strstr($value['text'], "\n"))
                                    {
                                        $this->text($value['text']);
                                    }
                                    else
                                    {
                                        $this->line($value['text']);
                                    }
                                }
                                else
                                {
                                    $this->addText($value['text']);
                                }
                            }
                        }
                        elseif ($this->canBeString($value))
                        {
                            // value can be converted to string
                            $this->addText((string) $value);
                        }
                        else
                        {
                            throw new Exception("Value of variable '".$Element['key']."' cannot be converted to a string");
                        }
                    }
                    elseif (in_array($Element['name'], array('h1', 'h2', 'h3', 'h4')))
                    {
                        // do nothing!
                    }
                    else
                    {
                        $this->addText(!$permitRawHtml ? self::escape($text, true) : $text);
                    }
                }
                else
                {
                    if (!isset($Element['comment']))
                    {
                        $this->addText(!$permitRawHtml ? self::escape($text, true) : $text);
                    }
                }
            }

            if ($hasName)
            {
                if ($Element['name'] == 'p')
                {
                    if ($this->pStyle)
                    {
                        $this->pStyle = null;
                    }
                    $this->textRun = null;
                }
                elseif ($Element['name'] == 'strong')
                {
                    $this->options['bold'] = false;
                }
                elseif ($Element['name'] == 'em')
                {
                    $this->options['italic'] = false;
                }
                elseif ($Element['name'] == 'u')
                {
                    $this->options['underline'] = 'none';
                }
                elseif (in_array($Element['name'], array('ul', 'ol')))
                {
                    $this->listStyleName = $parentListStyle;
                    $this->listDepth -= 1;
                }
                elseif ($Element['name'] == 'li')
                {
                    $this->listItemRun = null;
                }
                elseif ($Element['name'] == 'blockquote')
                {
                    if (isset($Element['attributes']['class']))
                    {
                        $this->pStyle = null;
                    }
                    else
                    {
                        $this->paragraphDepth -= 1;
                    }
                }
                elseif ($Element['name'] == 'table')
                {
                    $this->table = null;
                }
                elseif ($Element['name'] == 'tr')
                {
                    $this->row = null;
                    $this->currentColSpan = 0;
                }
                elseif (in_array($Element['name'], array('td', 'th')))
                {
                    if ($this->row)
                    {
                        if ($this->addTextToCell)
                        {
                            $this->cell = null;
                            $this->textRun = null;
                            $this->pStyleCell = null;
                            $this->addTextToCell = null;
                        }
                        
                        if ($this->currentColSpan > 0)
                        {
                            $this->currentColSpan -= 1;
                        }
                    }
                }
                elseif (in_array($Element['name'], array('thead', 'tbody')))
                {
                    $this->isTh = null;
                }
                elseif ($Element['name'] == 'fontstyle')
                {
                    $this->options['color'] = null;
                    $this->options['bgColor'] = null;
                    $this->options['size'] = null;
                }
            }
        }
        elseif ($hasName)
        {
            // self closing tags
        }
    }
    
    #
    # Convenient method for removing last PhpWord element in section
    #
    
    protected function cleanupTextRun()
    {
        if ($this->section && $this->textRun)
        {
            $lastIdx = $this->section->countElements() - 1;
            $this->section->removeElement($lastIdx);
            $this->textRun = null;
        }
    }
    
    #
    # Convenient method to get (and create if one doesn't exist) a vessel for text (textRun or Paragraph)
    #
    
    protected function getRun()
    {
        if ($this->outputMode == self::TYPE_DOCX)
        {
            $run = $this->listItemRun ? $this->listItemRun : $this->textRun;
            
            if (!$run)
            {
                if ($this->pStyle)
                {
                    $pStyle = $this->pStyle;
                }
                else
                {
                    switch ($this->paragraphDepth)
                    {
                        case 0: $pStyle = 'non_indented'; break;
                        case 1: $pStyle = 'indented'; break;
                        default: $pStyle = null;
                    }
                }
                
                $this->textRun = $this->section->addTextRun($pStyle);
                $run = $this->textRun;
            }
            
            return $run;
        }
        
        if ($this->outputMode == self::TYPE_ODT)
        {
            $run = $this->p;
            
            // create new Paragraph if it doesn't exist
            if (!$run) {
                if ($this->pStyle)
                {
                    $pStyle = $this->pStyle;
                }
                else
                {
                    switch ($this->paragraphDepth)
                    {
                        case 0: $pStyle = 'non_indented'; break;
                        case 1: $pStyle = 'indented'; break;
                        default: $pStyle = null;
                    }
                }
                
                $this->p = new Paragraph($this->getOdtParagraphStyle($pStyle));
                $run = $this->p;
            }
            
            return $run;
        }
        
        return null;
    }

    #
    # Convenient method to add text to current paragraph
    #

    protected function addText($text)
    {
        if ($text && $text != "\n")
        {
            if ($this->outputMode == self::TYPE_DOCX && (!$this->table || ($this->table && $this->addTextToCell)))
            {
                // handle ampersands, should be '&' without 'amp;'
                if (strpos($text, 'amp;') !== false) {
                    $text = str_replace('amp;', '', $text);
                }
                
                $run = $this->getRun();
                
                if ($this->elementIsBr($text))
                {
                    $run->addTextBreak();
                }
                else
                {
                    $run->addText($text, $this->options);
                }
            }
            elseif ($this->outputMode == self::TYPE_ODT)
            {
                // handle ampersands, should be '&amp;amp;'
                if (strpos($text, 'amp;') !== false) {
                    $text = str_replace('amp;', '', $text);
                }
                if (strpos($text, '&') !== false) {
                    $text = str_replace('&', '&amp;amp;', $text);
                }
                
                $run = $this->getRun();
                
                if ($this->elementIsBr($text))
                {
                    $run->addLineBreak();
                }
                else
                {
                    $run->addText($text, $this->wordOptionsToOdtTextStyle($this->options));
                }
            }
        }
    }

    protected function elementOdt(array $Element)
    {
        /*
        if ($this->safeMode)
        {
            $Element = $this->sanitiseElement($Element);
        }
        */

        # identity map if element has no handler
        $Element = $this->handle($Element);

        $hasName = isset($Element['name']);

        $parentList = null;
        $parentListStyle = null;

        if ($hasName)
        {
            if ($Element['name'] == 'p') {
                // do nothing!
                // Paragraph objects are made on demand in getRun method
            }
            elseif ($Element['name'] == 'strong')
            {
                $this->options['bold'] = true;
            }
            elseif ($Element['name'] == 'em')
            {
                $this->options['italic'] = true;
            }
            elseif ($Element['name'] == 'u')
            {
                $this->options['underline'] = 'single';
            }
            elseif ($Element['name'] == 'img')
            {
                if (isset($Element['attributes']['src']))
                {
                    $height = 0;
                    $width = 0;
                    $options = array();
                    if (isset($Element['attributes']['alt']) && $Element['attributes']['alt'])
                    {
                        $altArgs = explode(',',$Element['attributes']['alt']);
                        $countArgs = count($altArgs);
                        if ($countArgs >= 2)
                        {
                            $height = $altArgs[0];
                            $width = $altArgs[1];
                        }
                        elseif ($countArgs == 1)
                        {
                            $height = $altArgs[0];
                            $width = $altArgs[0];
                        }
                        $options['height'] = \PhpOffice\PhpWord\Shared\Converter::cmToPoint($height);
                        $options['width'] = \PhpOffice\PhpWord\Shared\Converter::cmToPoint($width);
                    }
                    
                    $run = $this->getRun();
                    $run->addImage($Element['attributes']['src'], "${width}cm", "${height}cm");
                }
            }
            elseif (in_array($Element['name'], array('h1', 'h2', 'h3', 'h4')))
            {
                $level = substr($Element['name'], -1);
                $pStyleName = 'h'.$this->headingsMap[$level];
                
                $p = new Paragraph($this->getOdtParagraphStyle($pStyleName));
                $p->addText($Element['text'], $this->getOdtTextStyle('bold'));
            }
            elseif (in_array($Element['name'], array('ol', 'ul')))
            {
                $parentListStyle = $this->listStyleName;
                $this->listDepth += 1;
                $this->listStyleName = $Element['name'] == 'ol' ? ($this->listIsRoman ? 'roman' : 'decimal') : ($this->listIsBullets ? 'bullet' : 'hyphen');

                if ($this->listDepth > 0)
                {
                    // mark list item as already added
                    $this->odtFinishedListItems[($this->listDepth-1).'_'.$this->list->getNumberOfElements()] = true;
                    $this->list->addParagraph($this->p);

                    $parentList = $this->list;
                }

                $this->list = new ODTList();
                if ($this->listDepth == 0)
                {
                    $this->list->setStyle($this->getOdtListStyle($this->listStyleName));
                }
            }
            elseif ($Element['name'] == 'li')
            {
                $this->p = new Paragraph();
            }
            elseif ($Element['name'] == 'hr')
            {
                new Paragraph();
            }
            elseif ($Element['name'] == 'br')
            {
                if ($this->p != null)
                {
                    $this->p->addLineBreak();
                }
                else
                {
                    new Paragraph();
                }
            }
            elseif ($Element['name'] == 'blockquote')
            {
                if (isset($Element['attributes']['class']))
                {
                    $this->pStyle = $Element['attributes']['class'];
                }
                else
                {
                    $this->paragraphDepth += 1;
                }                
            }
            elseif ($Element['name'] == 'table')
            {
                $tableStyle = $Element['table_styles'];
                $this->tableCellStylesMap = array();
                
                $rndString = self::randomString(5); // random 5 char string
                $this->table = new Table($rndString);
                $this->table->createColumns($Element['column_num']);
                
                $odtTableStyle = new TableStyle($this->table->getTableName());
                $odtTableStyle->setBorderModel(StyleConstants::COLLAPSING);
                
                if ($Element['are_widths_specified'])
                {
                    $total = 0;
                    for ($i = 0; $i < $Element['column_num']; ++$i)
                    {
                        if ($Element['column_widths'][$i])
                        {
                            $total += $Element['column_widths'][$i];
                            $this->table->getColumnStyle($i)->setWidth($Element['column_widths'][$i].'cm');
                        }
                    }

                    $alignment = StyleConstants::LEFT;
                    if (isset($tableStyle['alignment']))
                    {
                        if ($tableStyle['alignment'] === 'center')
                        {
                            $alignment = StyleConstants::CENTER;
                        }
                        elseif ($tableStyle['alignment'] === 'end')
                        {
                            $alignment = StyleConstants::RIGHT;
                        }
                    }
                    
                    $odtTableStyle->setWidth($total.'cm');
                    $odtTableStyle->setAlignment($alignment);
                }
                
                if ($Element['bg_color'])
                {
                    $odtTableStyle->setBgColor('#'.$Element['bg_color']);
                }
                
                $this->table->setStyle($odtTableStyle);

                $this->tableRows = array();
            }
            elseif ($Element['name'] == 'tr')
            {
                // Check if text of table headers exists.
                // If no text supplied, the table header row won't be created.
                $thHasText = false;
                if ($this->isTh && $Element['elements'])
                {
                    foreach ($Element['elements'] as $el)
                    {
                        if ($el['name'] == 'th' && trim($el['handler']['argument']) !== '')
                        {
                            $thHasText = true;
                            break;
                        }
                    }
                }

                if (!$this->isTh || $thHasText)
                {
                    $this->row = array();
                }
                
            }
            elseif (in_array($Element['name'], array('td', 'th')))
            {
                if (is_array($this->row))
                {
                    $cellStyle = array();
                    if ($Element['name'] === 'td')
                    {
                        $cellStyle = $Element['cell_styles'];
                    }
                    
                    // set cell alignment
                    $alignment = 'left';
                    if (isset($Element['column_alignment']) || isset($cellStyle['textAlignment']))
                    {
                        if (isset($cellStyle['textAlignment']))
                        {
                            $alignment = $cellStyle['textAlignment'];
                        }
                        else
                        {
                            // column alignment as fallback
                            $alignment = $Element['column_alignment'];
                        }
                    }
                    switch ($alignment)
                    {
                        case 'left': $this->pStyleCell = 'left_aligned'; break;
                        case 'right': $this->pStyleCell = 'right_aligned'; break;
                        case 'center': $this->pStyleCell = 'center_aligned'; break;
                    }
                    
                    // Save cell styles temporarily.
                    // They will be applied when table ends.
                    $cellStyleValues = array();
                    
                    $n = count($this->tableRows);
                    $m = count($this->row);
                    
                    // save cell background color
                    if (isset($cellStyle['bgColor']))
                    {
                        $cellStyleValues['bgColor'] = '#'.$cellStyle['bgColor'];
                    }
                    
                    // save vertical align
                    $valign = StyleConstants::TOP;
                    if (isset($cellStyle['valign']))
                    {
                        switch ($cellStyle['valign'])
                        {
                            case 'center': $valign = StyleConstants::CENTER; break;
                            case 'bottom': $valign = StyleConstants::BOTTOM; break;
                        }
                    }
                    $cellStyleValues['valign'] = $valign;
                    
                    $this->tableCellStylesMap[$n][$m] = $cellStyleValues;
                    $this->p = new Paragraph($this->getOdtParagraphStyle($this->pStyleCell));
                }
            }
            elseif ($Element['name'] == 'thead')
            {
                $this->isTh = true;
            }
            elseif ($Element['name'] == 'tbody')
            {
                $this->isTh = false;
            }
            elseif ($Element['name'] == 'fontstyle')
            {
                $color = $Element['color'];
                $bgColor = $Element['bgColor'];
                $size = $Element['size'];
                
                $this->options['color'] = $color;
                $this->options['bgColor'] = $bgColor;
                $this->options['size'] = $size;
                
                if (!isset($this->odtTextStyles[$this->odtBaseTextStyleNames[0].'_'.$color.'_'.$bgColor.'_'.$size]))
                {
                    foreach ($this->odtBaseTextStyleNames as $name)
                    {
                        $styleName = $name.'_'.$color.'_'.$bgColor.'_'.$size;
                        $textStyle = new TextStyle($styleName);
                        $textStyle->setFontName($this->fontName);
                        $textStyle->setFontSize($this->fontSize);

                        if (strpos($name, 'bold') !== false)
                        {
                            $textStyle->setBold();
                        }
                        if (strpos($name, 'italic') !== false)
                        {
                            $textStyle->setItalic();
                        }
                        if (strpos($name, 'underline') !== false)
                        {
                            $textStyle->setTextUnderline();
                        }

                        if ($color)
                        {
                            $colorOdt = (ctype_xdigit($color) ? '#' : '').$color;
                            $textStyle->setColor($colorOdt);
                        }
                        if ($bgColor)
                        {
                            $colorOdt = (ctype_xdigit($bgColor) ? '#' : '').$bgColor;
                            $textStyle->setTextBackgroundColor($colorOdt);
                        }
                        if ($size)
                        {
                            $textStyle->setFontSize($size);
                        }
                        
                        $this->odtTextStyles[$styleName] = $textStyle;
                    }
                }
                
            }
        }

        $permitRawHtml = false;

        if (isset($Element['text']))
        {
            $text = $Element['text'];
        }

        
        // very strongly consider an alternative if you're writing an
        // extension
        elseif (isset($Element['rawHtml']))
        {
            $text = $Element['rawHtml'];

            $allowRawHtmlInSafeMode = isset($Element['allowRawHtmlInSafeMode']) && $Element['allowRawHtmlInSafeMode'];
            $permitRawHtml = !$this->safeMode || $allowRawHtmlInSafeMode;
        }
        

        $hasContent = isset($text) || isset($Element['element']) || isset($Element['elements']);

        if ($hasContent)
        {
            if (isset($Element['elements']))
            {
                $this->elementsOdt($Element['elements']);
            }
            elseif (isset($Element['element']))
            {
                $this->elementOdt($Element['element']);
            }
            else
            {
                if ($hasName)
                {
                    if ($Element['name'] == 'var')
                    {
                        // handle variable value
                        
                        $value = $this->varConverter->evaluate($Element['key'], $Element['params']);
                        
                        if (is_array($value) && isset($value['markdown'], $value['text']))
                        {
                            // value is array
                            
                            if ($value['text'] != '')
                            {
                                if ($value['markdown'])
                                {
                                    // value is markdown text, so parse it and handle elements
                                    if (strstr($value['text'], "\n"))
                                    {
                                        $this->text($value['text']);
                                    }
                                    else
                                    {
                                        $this->line($value['text']);
                                    }
                                }
                                else
                                {
                                    $this->addText($value['text']);
                                }
                            }
                        }
                        elseif ($this->canBeString($value))
                        {
                            // value can be converted to string
                            $this->addText((string) $value);
                        }
                        else
                        {
                            throw new Exception("Value of variable '".$Element['key']."' cannot be converted to a string");
                        }
                    }
                    elseif (in_array($Element['name'], array('h1', 'h2', 'h3', 'h4')))
                    {
                        // do nothing!
                    }
                    else
                    {
                        $this->addText(!$permitRawHtml ? self::escape($text, true) : $text);
                    }
                }
                else
                {
                    if (!isset($Element['comment']))
                    {
                        $this->addText(!$permitRawHtml ? self::escape($text, true) : $text);
                    }
                }
            }

            if ($hasName)
            {
                if ($Element['name'] == 'p')
                {
                    if ($this->pStyle)
                    {
                        $this->pStyle = null;
                    }
                    $this->p = null;
                }
                elseif ($Element['name'] == 'strong')
                {
                    $this->options['bold'] = false;
                }
                elseif ($Element['name'] == 'em')
                {
                    $this->options['italic'] = false;
                }
                elseif ($Element['name'] == 'u')
                {
                    $this->options['underline'] = 'none';
                }
                elseif (in_array($Element['name'], array('ul', 'ol')))
                {
                    if ($parentList && $this->listDepth > 0)
                    {
                        $parentList->addSubList($this->list);
                        $this->list = $parentList;
                    } else {
                        $this->list = null;
                    }
                    $this->listStyleName = $parentListStyle;
                    $this->listDepth -= 1;
                }
                elseif ($Element['name'] == 'li')
                {
                    // create unique key for item: <list depth>_<item_position>
                    $key = $this->listDepth.'_'.($this->list->getNumberOfElements()-1);

                    // skip list item if already added before
                    if (!isset($this->odtFinishedListItems[$key]))
                    {
                        $this->list->addParagraph($this->p);
                    }
                    else
                    {
                        // mark as handled by unsetting map value
                        unset($this->odtFinishedListItems[$key]);
                    }
                    $this->p = null;
                }
                elseif ($Element['name'] == 'blockquote')
                {
                    if (isset($Element['attributes']['class']))
                    {
                        $this->pStyle = null;
                    }
                    else
                    {
                        $this->paragraphDepth -= 1;
                    }
                }
                elseif ($Element['name'] == 'table')
                {
                    $this->table->addRows($this->tableRows);
                    
                    // After adding rows to the table, cell styles are automatically created.
                    // Iterate through the cell styles and set values accordingly.                    
                    for ($i = 0; $i < $Element['row_num']; ++$i)
                    {
                        for ($j = 0; $j < $Element['column_num']; ++$j)
                        {
                            $cellStyle = $this->table->getCellStyle($j, $i);
                            
                            // set cell border
                            if (isset($Element['table_styles']['borderSize']))
                            {
                                $borderColor = '000000';
                                if (isset($Element['table_styles']['borderColor']))
                                {
                                    $borderColor = $Element['table_styles']['borderColor'];
                                }
                                
                                $borderSizePt = $Element['table_styles']['borderSize'] / 20.0; // twip -> pt
                                $cellStyle->setBorder('#'.$borderColor, StyleConstants::SOLID, $borderSizePt.'pt');
                            }
                            
                            // set cell padding (even though it's called margin)
                            if (isset($Element['table_styles']['cellMargin']))
                            {
                                $cellStyle->setPadding($Element['table_styles']['cellMargin'].'cm');
                            }
                            
                            $cellValues = $this->tableCellStylesMap[$i][$j];
                            
                            // set cell background color
                            if (isset($cellValues['bgColor']))
                            {
                                $cellStyle->setBgColor($cellValues['bgColor']);
                            }
                            
                            // set vertical align
                            $cellStyle->setVerticalAlign($cellValues['valign']);
                        }
                    }
                    
                    $this->tableCellStylesMap = null;
                    $this->tableRows = null;
                    $this->table = null;
                }
                elseif ($Element['name'] == 'tr')
                {
                    if (is_array($this->row))
                    {
                        $this->tableRows[] = $this->row;
                    }
                    
                    $this->row = null;
                }
                elseif (in_array($Element['name'], array('td', 'th')))
                {
                    $this->pStyleCell = null;

                    if (is_array($this->row))
                    {
                        $this->row[] = $this->p;
                    }

                    $this->p = null;
                }
                elseif (in_array($Element['name'], array('thead', 'tbody')))
                {
                    $this->isTh = null;
                }
                elseif ($Element['name'] == 'fontstyle')
                {
                    $this->options['color'] = null;
                    $this->options['bgColor'] = null;
                    $this->options['size'] = null;
                }
            }
        }
        elseif ($hasName)
        {
            // self closing tags
        }
    }

    protected function elements(array $Elements)
    {
        $autoBreak = true;

        foreach ($Elements as $Element)
        {
            if (empty($Element))
            {
                continue;
            }

            $autoBreakNext = (isset($Element['autobreak'])
                ? $Element['autobreak'] : isset($Element['name'])
            );
            // (autobreak === false) covers both sides of an element
            $autoBreak = !$autoBreak ? $autoBreak : $autoBreakNext;

            $this->markup .= ($autoBreak ? "\n" : '');
            $this->element($Element);

            $autoBreak = $autoBreakNext;
        }

        $this->markup .= $autoBreak ? "\n" : '';
    }

    protected function elementsDocx(array $Elements)
    {
        $autoBreak = true;

        foreach ($Elements as $Element)
        {
            if (empty($Element))
            {
                continue;
            }

            $autoBreakNext = (isset($Element['autobreak'])
                ? $Element['autobreak'] : isset($Element['name'])
            );
            // (autobreak === false) covers both sides of an element
            $autoBreak = !$autoBreak ? $autoBreak : $autoBreakNext;

            // $markup .= ($autoBreak ? "\n" : '') . $this->elementDocx($Element);
            $this->elementDocx($Element);
            $autoBreak = $autoBreakNext;
        }
        // $markup .= $autoBreak ? "\n" : '';
    }
    
    protected function elementsOdt(array $Elements)
    {
        $autoBreak = true;

        foreach ($Elements as $Element)
        {
            if (empty($Element))
            {
                continue;
            }

            $autoBreakNext = (isset($Element['autobreak'])
                ? $Element['autobreak'] : isset($Element['name'])
            );
            // (autobreak === false) covers both sides of an element
            $autoBreak = !$autoBreak ? $autoBreak : $autoBreakNext;

            // $markup .= ($autoBreak ? "\n" : '') . $this->elementOdt($Element);
            $this->elementOdt($Element);
            $autoBreak = $autoBreakNext;
        }
        // $markup .= $autoBreak ? "\n" : '';
    }

    # ~

    protected function li($lines)
    {
        $Elements = $this->linesElements($lines);

        if ( ! in_array('', $lines)
            and isset($Elements[0]) and isset($Elements[0]['name'])
            and $Elements[0]['name'] === 'p'
        ) {
            unset($Elements[0]['name']);
        }

        return $Elements;
    }

    #
    # AST Convenience
    #

    /**
     * Replace occurrences $regexp with $Elements in $text. Return an array of
     * elements representing the replacement.
     */
    protected static function pregReplaceElements($regexp, $Elements, $text)
    {
        $newElements = array();

        while (preg_match($regexp, $text, $matches, PREG_OFFSET_CAPTURE))
        {
            $offset = $matches[0][1];
            $before = substr($text, 0, $offset);
            $after = substr($text, $offset + strlen($matches[0][0]));

            $newElements[] = array('text' => $before);

            foreach ($Elements as $Element)
            {
                $newElements[] = $Element;
            }

            $text = $after;
        }

        $newElements[] = array('text' => $text);

        return $newElements;
    }
    
    #
    # Helper method to determine if a variable can be transformed into a string
    #
    
    protected function canBeString($value)
    {
        return $value === null || is_scalar($value) || (is_object($value) && method_exists($value, '__toString'));
    }

    #
    # Replaces all occurrences of variables in a string
    # with placeholders {0}, {1}, ..., one after the other.
    # This is sometimes convenient as to disallow variable
    # parameter string characters to interfere with MD parsing.
    #
    # Function argument $text is passed by reference.
    #
    # Return value is a map of key-value pairs, each representing
    # a placeholder and variable text, respectively.
    #
    
    protected function replaceVarsWithPlaceholders(&$text)
    {
        $vars = array();
        if (preg_match_all($this->VariableRegexAnywhere, $text, $varMatches)) {
            foreach ($varMatches[0] as $i => $varText) {
                $key = '{'.$i.'}';
                $vars[$key] = $varText;
                $text = str_replace($varText, $key, $text);
            }
        }
        return $vars;
    }
    
    #
    # Replaces all occurrences of placeholders {0}, {1}, ...,
    # in a string with variables from the given map.
    #
    # Function argument $text is passed by reference.
    #
    
    protected function replacePlaceholdersWithVars(&$text, array $vars)
    {
        foreach ($vars as $key => $varText) {
            $text = str_replace($key, $varText, $text);
        }
    }
    
    #
    # Converts the options array used in PhpWord to an ODT text style object
    #

    protected function wordOptionsToOdtTextStyle(array $options)
    {
        $isBold = isset($options['bold']) && $options['bold'];
        $isItalic = isset($options['italic']) && $options['italic'];
        $isUnderline = isset($options['underline']) && $options['underline'] === 'single';
        $hasFontOptions = isset($options['color']) || isset($options['bgColor']) || isset($options['size']);

        $style = 'default';

        if ($isBold && $isItalic && $isUnderline && isset($this->odtTextStyles['bold_italic_underline']))
        {
            $style = 'bold_italic_underline';
        }
        elseif ($isBold && $isItalic && isset($this->odtTextStyles['bold_italic']))
        {
            $style = 'bold_italic';
        }
        elseif ($isBold && $isUnderline && isset($this->odtTextStyles['bold_underline']))
        {
            $style = 'bold_underline';
        }
        elseif ($isItalic && $isUnderline && isset($this->odtTextStyles['italic_underline']))
        {
            $style = 'italic_underline';
        }
        elseif ($isBold && isset($this->odtTextStyles['bold']))
        {
            $style = 'bold';
        }
        elseif ($isItalic && isset($this->odtTextStyles['italic']))
        {
            $style = 'italic';
        }
        elseif ($isUnderline && isset($this->odtTextStyles['underline']))
        {
            $style = 'underline';
        }

        if ($hasFontOptions)
        {
            // append color, background color and size to style name
            $style .= '_'.$options['color'].'_'.$options['bgColor'].'_'.$options['size'];
        }

        return isset($this->odtTextStyles[$style]) ? $this->odtTextStyles[$style] : null;
    }

    #
    # Gets ODT list style based on style name
    #

    protected function getOdtTextStyle($key)
    {
        return isset($this->odtTextStyles[$key]) ? $this->odtTextStyles[$key] : null;
    }

    #
    # Gets ODT text style based on style name
    #

    protected function getOdtListStyle($key)
    {
        return isset($this->odtListStyles[$key]) ? $this->odtListStyles[$key] : null;
    }

    #
    # Gets ODT paragraph style based on style name
    #

    protected function getOdtParagraphStyle($key)
    {
        return isset($this->odtParagraphStyles[$key]) ? $this->odtParagraphStyles[$key] : null;
    }
    
    #
    # Adds new ODT paragraph style to Parsedown
    #
    
    public function addOdtParagraphStyle($pStyle)
    {
        $this->odtParagraphStyles[$pStyle->getStyleName()] = $pStyle;
    }
    
    #
    # Parses text to extract emphasis (bold, italic, underline) information
    # which is stored in a boolean map and returns underlying word.
    # Example:
    # Input: '*_foo_*'
    # Output: returns 'foo', $emMap = ['em' => true, 'u' => true]
    #

    public function extractEmphasisFromSource($text, array &$emMap)
    {
        $le = $this->lineElements($text);
        $innerText = $text;
        while (isset($le[1]['name']))
        {
            $el = $le[1];
            if (in_array($el['name'], array('em', 'strong', 'u')) && isset($el['handler']['argument']))
            {
                $key = $el['name'];
            }
            else
            {
                break;
            }

            $emMap[$key] = true;
            $innerText = $el['handler']['argument'];
            $le = $this->lineElements($innerText);
        }
        
        return $innerText;
    }

    #
    # Returns a Markdown text representing an encapsulated string value based on an emphasis array map
    # Example:
    # Input: $str = 'foo', $emMap = ['em' => true, 'u' => true]
    # Output: returns '*_foo_*'
    #

    public function insertEmphasisToSource($str, array $emMap)
    {
        $text = '';
        foreach ($emMap as $key => $val)
        {
            $text .= $key == 'em' && $val ? '*' : ($key == 'strong' && $val ? '**' : ($key == 'u' && $val ? '_' : ''));
        }
        $text .= $str;
        $emsReverse = array_reverse($emMap);
        foreach ($emsReverse as $key => $val) {
            $text .= $key == 'em' && $val ? '*' : ($key == 'strong' && $val ? '**' : ($key == 'u' && $val ? '_' : ''));
        }
        return $text;
    }
    
    #
    # Checks if supplied markup is a <br> tag
    #

    protected function elementIsBr($markup)
    {
        return preg_match($this->BrRegex, $markup);
    }

    #
    # Deprecated Methods
    #

    protected function sanitiseElement(array $Element)
    {
        static $goodAttribute = '/^[a-zA-Z0-9][a-zA-Z0-9-_]*+$/';
        static $safeUrlNameToAtt  = array(
            'a'   => 'href',
            'img' => 'src',
        );

        if ( ! isset($Element['name']))
        {
            unset($Element['attributes']);
            return $Element;
        }

        if (isset($safeUrlNameToAtt[$Element['name']]))
        {
            $Element = $this->filterUnsafeUrlInAttribute($Element, $safeUrlNameToAtt[$Element['name']]);
        }

        if ( ! empty($Element['attributes']))
        {
            foreach ($Element['attributes'] as $att => $val)
            {
                # filter out badly parsed attribute
                if ( ! preg_match($goodAttribute, $att))
                {
                    unset($Element['attributes'][$att]);
                }
                # dump onevent attribute
                elseif (self::striAtStart($att, 'on'))
                {
                    unset($Element['attributes'][$att]);
                }
            }
        }

        return $Element;
    }

    protected function filterUnsafeUrlInAttribute(array $Element, $attribute)
    {
        foreach ($this->safeLinksWhitelist as $scheme)
        {
            if (self::striAtStart($Element['attributes'][$attribute], $scheme))
            {
                return $Element;
            }
        }

        $Element['attributes'][$attribute] = str_replace(':', '%3A', $Element['attributes'][$attribute]);

        return $Element;
    }

    #
    # Static Methods
    #

    protected static function escape($text, $allowQuotes = false)
    {
        return htmlspecialchars($text, $allowQuotes ? ENT_NOQUOTES : ENT_QUOTES, 'UTF-8');
    }

    protected static function striAtStart($string, $needle)
    {
        $len = strlen($needle);

        if ($len > strlen($string))
        {
            return false;
        }
        else
        {
            return strtolower(substr($string, 0, $len)) === strtolower($needle);
        }
    }

    static function instance($name = 'default')
    {
        if (isset(self::$instances[$name]))
        {
            return self::$instances[$name];
        }

        $instance = new static();

        self::$instances[$name] = $instance;

        return $instance;
    }
    
    public static function randomString($length)
    {
        return substr(md5(microtime()), rand(0, 26), $length); // random char string
    }

    private static $instances = array();

    #
    # Fields
    #

    protected $DefinitionData;

    #
    # Read-Only

    protected $specialCharacters = array(
        '\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '>', '#', '+', '-', '.', '!', '|', '~',
    );

    protected $StrongRegex = array(
        '*' => '/^[*]{2}((?:\\\\\*|[^*]|[*][^*]*+[*])+?)[*]{2}(?![*])/s',
    );

    protected $EmRegex = array(
        '*' => '/^[*]((?:\\\\\*|[^*]|[*][*][^*]+?[*][*])+?)[*](?![*])/s',
    );

    protected $UnderlineRegex = '/^_((?:\\\\_|[^_])+?)_(?!_)\b/us';

    protected $fontStyleRegex = '/^`\[(?:([\dA-Fa-f]{6})?(?::([\dA-Fa-f]{6}))?(?:\+(\d+(?:\.\d+)?))?)\]((?:\\\\`|[^`])+?)`/us';
    
    protected $VariableRegex = '/^\$\{((?:\\\\\$|[^$])+?)\}/us';
    
    protected $VariableRegexAnywhere = '/\$\{((?:\\\\\$|[^$])+?)\}/us'; // same as $VariableRegex except for ^

    protected $BrRegex = "/^<br\W*?\/?>$/";
    
    protected $StyleOptionsRegexAnywhere = '/\[((?:\s*\w+\s*\=\s*[a-zA-Z0-9.]+\s*(?:,\s*\w+\s*\=\s*[a-zA-Z0-9.]+\s*)*)+?)\]/us';

    protected $regexHtmlAttribute = '[a-zA-Z_:][\w:.-]*+(?:\s*+=\s*+(?:[^"\'=<>`\s]+|"[^"]*+"|\'[^\']*+\'))?+';

    protected $voidElements = array(
        'area', 'base', 'br', 'col', 'command', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source',
    );

    protected $textLevelElements = array(
        'a', 'br', 'bdo', 'abbr', 'blink', 'nextid', 'acronym', 'basefont',
        'b', 'em', 'big', 'cite', 'small', 'spacer', 'listing',
        'i', 'rp', 'del', 'code',          'strike', 'marquee',
        'q', 'rt', 'ins', 'font',          'strong',
        's', 'tt', 'kbd', 'mark',
        'u', 'xm', 'sub', 'nobr',
                   'sup', 'ruby',
                   'var', 'span',
                   'wbr', 'time',
    );
    
    protected $availableTableStyles = array(
        'bgColor',
        'borderColor', /*'borderTopColor', 'borderBottomColor', 'borderRightColor', 'borderLeftColor',*/
        'borderSize', /*'borderTopSize', 'borderBottomSize', 'borderRightSize', 'borderLeftSize',*/
        'cellMargin', /*'cellMarginTop', 'cellMarginBottom', 'cellMarginRight', 'cellMarginLeft',*/
        'alignment',
    );
    
    protected $availableCellStyles = array(
        'gridSpan', 'bgColor', 'textAlignment', 'valign',
    );
}
