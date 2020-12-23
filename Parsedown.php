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

    const TYPE_HTML = 'html';
    const TYPE_DOCX = 'docx';
    const TYPE_ODT = 'odt';

    # ~

    protected $outputMode; // html, docx, odt

    protected $varConverter;

    protected $fontSize;
    protected $options = [];

    protected $phpWord;
    protected $section;
    protected $textRun;
    protected $listStyleName;
    protected $lastListStyleName;
    protected $listDepth;

    protected $odt;
    protected $odtTextStyles;
    protected $odtParagraphStyles;
    protected $odtListStyles;
    protected $p;
    protected $list;
    protected $odtFinishedListItems;

    public function __construct($mode = null)
    {
        if ($mode)
        {
            $this->outputMode = $mode;
            $this->fontSize = 11;
            $this->listDepth = -1;
    
            if ($this->outputMode == self::TYPE_ODT)
            {
                $this->initializeOdtParameters();
            }
            elseif ($this->outputMode == self::TYPE_DOCX)
            {
                $this->initializePhpWordParameters();
            }
        }
    }

    protected function initializeOdtParameters()
    {
        $this->odt = ODT::getInstance();
        $this->odtTextStyles = [];
        $this->odtParagraphStyles = [];
        $this->odtListStyles = [];
        $this->odtFinishedListItems = [];
        
        $boldTextStyle = new TextStyle('bold');
        $boldTextStyle->setBold();
        $boldTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$boldTextStyle->getStyleName()] = $boldTextStyle;
        
        $italicTextStyle = new TextStyle('italic');
        $italicTextStyle->setItalic();
        $italicTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$italicTextStyle->getStyleName()] = $italicTextStyle;

        $underlineTextStyle = new TextStyle('underline');
        $underlineTextStyle->setTextUnderline();
        $underlineTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$underlineTextStyle->getStyleName()] = $underlineTextStyle;
        
        $boldItalicTextStyle = new TextStyle('bold_italic');
        $boldItalicTextStyle->setBold();
        $boldItalicTextStyle->setItalic();
        $boldItalicTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$boldItalicTextStyle->getStyleName()] = $boldItalicTextStyle;

        $boldUnderlineTextStyle = new TextStyle('bold_underline');
        $boldUnderlineTextStyle->setBold();
        $boldUnderlineTextStyle->setTextUnderline();
        $boldUnderlineTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$boldUnderlineTextStyle->getStyleName()] = $boldUnderlineTextStyle;

        $italicUnderlineTextStyle = new TextStyle('italic_underline');
        $italicUnderlineTextStyle->setItalic();
        $italicUnderlineTextStyle->setTextUnderline();
        $italicUnderlineTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$italicUnderlineTextStyle->getStyleName()] = $italicUnderlineTextStyle;
        
        $boldItalicUnderlineTextStyle = new TextStyle('bold_italic_underline');
        $boldItalicUnderlineTextStyle->setBold();
        $boldItalicUnderlineTextStyle->setItalic();
        $boldItalicUnderlineTextStyle->setTextUnderline();
        $boldItalicUnderlineTextStyle->setFontSize($this->fontSize);
        $this->odtTextStyles[$boldItalicUnderlineTextStyle->getStyleName()] = $boldItalicUnderlineTextStyle;

        $heading1 = new ParagraphStyle('h1');
        $heading1->setTextAlign(StyleConstants::CENTER);
        $heading1->setVerticalMargin('0.2cm', '0.2cm');
        $this->odtParagraphStyles[$heading1->getStyleName()] = $heading1;

        $heading2 = new ParagraphStyle('h2');
        $heading2->setTextAlign(StyleConstants::CENTER);
        $heading2->setVerticalMargin('0.2cm', '0cm');
        $this->odtParagraphStyles[$heading2->getStyleName()] = $heading2;

        $heading3 = new ParagraphStyle('h3');
        $heading3->setTextAlign(StyleConstants::CENTER);
        $heading3->setVerticalMargin('0cm', '0cm');
        $this->odtParagraphStyles[$heading3->getStyleName()] = $heading3;

        $heading4 = new ParagraphStyle('h4');
        $heading4->setTextAlign(StyleConstants::CENTER);
        $heading4->setVerticalMargin('0cm', '0.2cm');
        $this->odtParagraphStyles[$heading4->getStyleName()] = $heading4;

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
    }

    protected function initializePhpWordParameters()
    {
        \PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);
        $this->phpWord = new \PhpOffice\PhpWord\PhpWord();
        $this->phpWord->setDefaultFontName('Times New Roman');
        $this->phpWord->setDefaultFontSize($this->fontSize);
        $this->phpWord->setDefaultParagraphStyle(['align' => 'left', 'spaceAfter' => 0, 'spaceBefore' => 0]);
        $this->phpWord->addTitleStyle(1, ['bold' => true], ['spaceAfter' => 200, 'spaceBefore' => 200, 'align' => 'center']);
        $this->phpWord->addTitleStyle(11, ['bold' => true], ['spaceAfter' => 0, 'spaceBefore' => 200, 'align' => 'center']);
        $this->phpWord->addTitleStyle(12, ['bold' => true], ['spaceAfter' => 0, 'spaceBefore' => 0, 'align' => 'center']);
        $this->phpWord->addTitleStyle(13, ['bold' => true], ['spaceAfter' => 200, 'spaceBefore' => 0, 'align' => 'center']);
        $this->phpWord->addNumberingStyle('roman', ['type' => 'multilevel','levels' => [['format' => 'upperRoman', 'text' => '%1', 'left' => 720, 'hanging' => 720, 'tabPos' => 720], ['format' => 'lowerRoman', 'text' => '%2', 'left' => 1000, 'hanging' => 720, 'tabPos' => 1000]]]);
        $this->phpWord->addNumberingStyle('decimal', ['type' => 'multilevel','levels' => [['format' => 'decimal', 'text' => '%1.', 'left' => 720, 'hanging' => 720, 'tabPos' => 720], ['format' => 'decimal', 'text' => '%2.', 'left' => 1000, 'hanging' => 720, 'tabPos' => 1000]]]);
        $this->phpWord->addNumberingStyle('bullet', ['type' => 'multilevel', 'levels' => [['format' => 'bullet', 'text' => '-', 'left' => 520, 'hanging' => 200, 'tabPos' => 520], ['format' => 'bullet', 'text' => '-', 'left' => 880, 'hanging' => 200, 'tabPos' => 880], ['format' => 'bullet', 'text' => '-', 'left' => 1080, 'hanging' => 200, 'tabPos' => 1080]]]);
    
        $this->section = $this->phpWord->addSection([
            'marginTop' => 600,
            'marginBottom' => 600,
            'marginRight' => 800,
            'marginLeft' => 800]);
    }

    public function getVarConverter()
    {
        return $this->varConverter;
    }

    public function setVarConverter($varConverter)
    {
        $this->varConverter = $varConverter;
    }

    public function getPhpWord()
    {
        return $this->phpWord;
    }

    public function setPhpWord(\PhpOffice\PhpWord\PhpWord $phpWord)
    {
        $this->phpWord = $phpWord;
    }

    function getTree($text)
    {
        return $this->textElements($text);
    }

    function getHtmlFromTree(array $elements)
    {
        return trim($this->elements($elements), "\n");
    }

    function getDocxFromTree(array $elements)
    {
        return $this->elementsDocx($elements);
    }

    function getOdtFromTree(array $elements)
    {
        return $this->elementsOdt($elements);
    }

    function text($text)
    {
        # parse
        $Elements = $this->textElements($text);

        # convert
        if ($this->outputMode == self::TYPE_ODT)
        {
            $markup = $this->elementsOdt($Elements);
        }
        elseif ($this->outputMode == self::TYPE_DOCX)
        {
            $markup = $this->elementsDocx($Elements);
        }
        elseif ($this->outputMode == self::TYPE_HTML)
        {
            $markup = $this->elements($Elements);
        }
        else
        {
            throw new Exception('Unknown document type');
        }

        # trim line breaks
        $markup = trim($markup, "\n");

        return $markup;
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

    protected $breaksEnabled;

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
        '*' => array('Rule', 'List'),
        '+' => array('List'),
        '-' => array('SetextHeader', 'Table', 'Rule', 'List'),
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
        '=' => array('SetextHeader'),
        '>' => array('Quote'),
        '[' => array('Reference'),
        '_' => array('Rule'),
        '`' => array('FencedCode'),
        '|' => array('Table'),
        '~' => array('FencedCode'),
    );

    # ~

    protected $unmarkedBlockTypes = array(
        'Code',
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

            // TODO: REMOVE
            //var_dump('LINE:',$Line);
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

        
                // TODO: REMOVE
                //var_dump('linesElements:',$Elements);

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
                'htext' => $text,
            ),
        );

        return $Block;
    }

    #
    # List

    protected function blockList($Line, array $CurrentBlock = null)
    {
        list($name, $pattern) = $Line['text'][0] <= '-' ? array('ul', '[*+-]') : array('ol', '[0-9]{1,9}+[.\)]');

        // TODO: REMOVE
        //var_dump('LIST:',$Line);

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

        if ($Line['text'][0] === '[' and $this->blockReference($Line))
        {
            return $Block;
        }

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

        if (chop($Line['text'], ' -:|') !== '')
        {
            return;
        }

        $alignments = array();

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

            if ($dividerCell[0] === ':')
            {
                $alignment = 'left';
            }

            if (substr($dividerCell, - 1) === ':')
            {
                $alignment = $alignment === 'left' ? 'center' : 'right';
            }

            $alignments []= $alignment;
        }

        # ~

        $HeaderElements = array();

        $header = $Block['element']['handler']['argument'];

        $header = trim($header);
        $header = trim($header, '|');

        $headerCells = explode('|', $header);

        if (count($headerCells) !== count($alignments))
        {
            return;
        }

        foreach ($headerCells as $index => $headerCell)
        {
            $headerCell = trim($headerCell);

            $HeaderElement = array(
                'name' => 'th',
                'handler' => array(
                    'function' => 'lineElements',
                    'argument' => $headerCell,
                    'destination' => 'elements',
                )
            );

            if (isset($alignments[$index]))
            {
                $alignment = $alignments[$index];

                $HeaderElement['attributes'] = array(
                    'style' => "text-align: $alignment;",
                );
            }

            $HeaderElements []= $HeaderElement;
        }

        # ~

        $Block = array(
            'alignments' => $alignments,
            'identified' => true,
            'element' => array(
                'name' => 'table',
                'elements' => array(),
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

        if (count($Block['alignments']) === 1 or $Line['text'][0] === '|' or strpos($Line['text'], '|'))
        {
            $Elements = array();

            $row = $Line['text'];

            $row = trim($row);
            $row = trim($row, '|');

            preg_match_all('/(?:(\\\\[|])|[^|`]|`[^`]++`|`)++/', $row, $matches);

            $cells = array_slice($matches[0], 0, count($Block['alignments']));

            foreach ($cells as $index => $cell)
            {
                $cell = trim($cell);

                $Element = array(
                    'name' => 'td',
                    'handler' => array(
                        'function' => 'lineElements',
                        'argument' => $cell,
                        'destination' => 'elements',
                    )
                );

                if (isset($Block['alignments'][$index]))
                {
                    $Element['attributes'] = array(
                        'style' => 'text-align: ' . $Block['alignments'][$index] . ';',
                    );
                }

                $Elements []= $Element;
            }

            $Element = array(
                'name' => 'tr',
                'elements' => $Elements,
            );

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
        '`' => array('Code'),
        '~' => array('Strikethrough'),
        '\\' => array('EscapeSequence'),
        '$' => array('Var'),
    );

    # ~

    protected $inlineMarkerList = '!*_&[:<`~\\$';

    #
    # ~
    #

    public function line($text, $nonNestables = array())
    {
        return $this->elements($this->lineElements($text, $nonNestables));
    }
    
    public function lineDocx($text, $nonNestables = array())
    {
        return $this->elementsDocx($this->lineElements($text, $nonNestables));
    }
    
    public function lineOdt($text, $nonNestables = array())
    {
        return $this->elementsOdt($this->lineElements($text, $nonNestables));
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

        //if ($Excerpt['text'][1] === $marker and preg_match($this->StrongRegex[$marker], $Excerpt['text'], $matches))
        if ($Excerpt['text'][1] === $marker and preg_match($this->StrongRegex['*'], $Excerpt['text'], $matches))
        {
            $emphasis = 'strong';
        }
        //elseif (preg_match($this->EmRegex[$marker], $Excerpt['text'], $matches))
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
                'text' => $varPattern,
                /*
                'handler' => array(
                    'function' => 'lineText',
                    //'function' => 'varDefinition',
                    'argument' => $varPattern,
                    'destination' => 'elements',
                ),
                */
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

            $Element[$destination] = $this->{$function}($argument, $Element['nonNestables']);

            if ($destination === 'handler')
            {
                $Element = $this->handle($Element);
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

        $markup = '';

        if ($hasName)
        {
            $markup .= '<' . $Element['name'];

            if (isset($Element['attributes']))
            {
                foreach ($Element['attributes'] as $name => $value)
                {
                    if ($value === null)
                    {
                        continue;
                    }

                    $markup .= " $name=\"".self::escape($value).'"';
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
            $markup .= $hasName ? '>' : '';

            if (isset($Element['elements']))
            {
                $markup .= $this->elements($Element['elements']);
            }
            elseif (isset($Element['element']))
            {
                $markup .= $this->element($Element['element']);
            }
            else
            {
                if ($hasName && $Element['name'] == 'var')
                {
                    $value = $this->varConverter->evaluate($Element['key'], $Element['params']);
                    if (is_array($value) && isset($value['markdown'], $value['text']))
                    {
                        // TODO: check existence of $tree[0]['handler']['argument']
                        $tree = $this->getTree($value['text']);
                        $markup .= $this->elements($this->lineElements($tree[0]['handler']['argument']));
                    }
                    else
                    {
                        $markup .= $value;
                    }
                    
                }
                else
                {
                    if (!$permitRawHtml)
                    {
                        $markup .= self::escape($text, true);
                    }
                    else
                    {
                        $markup .= $text;
                    }
                }
            }

            $markup .= $hasName ? '</' . $Element['name'] . '>' : '';
        }
        elseif ($hasName)
        {
            $markup .= ' />';
        }

        return $markup;
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

        $markup = '';

        if ($hasName)
        {

            if ($Element['name'] == 'p') {
                $this->textRun = $this->section->addTextRun();
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
                    $options = [];
                    if (isset($Element['attributes']['alt']) && $Element['attributes']['alt'])
                    {
                        $altArgs = explode(',',$Element['attributes']['alt']);
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
                    $this->section->addImage($Element['attributes']['src'], $options);
                }
            }
            elseif ($Element['name'] == 'h1')
            {
                $this->section->addTitle($Element['htext'], 1);
            }
            elseif ($Element['name'] == 'h2')
            {
                $this->section->addTitle($Element['htext'], 11);
            }
            elseif ($Element['name'] == 'h3')
            {
                $this->section->addTitle($Element['htext'], 12);
            }
            elseif ($Element['name'] == 'h4')
            {
                $this->section->addTitle($Element['htext'], 13);
            }
            elseif ($Element['name'] == 'ol')
            {
                $parentListStyle = $this->listStyleName;
                $this->listDepth += 1;
                $this->listStyleName = 'roman';
            }
            elseif ($Element['name'] == 'ul')
            {
                $parentListStyle = $this->listStyleName;
                $this->listDepth += 1;
                $this->listStyleName = 'bullet';
            }
            elseif ($Element['name'] == 'li')
            {
                $this->textRun = $this->section->addListItemRun($this->listDepth, $this->listStyleName);
            }

            /*
            $markup .= '<' . $Element['name'];

            if (isset($Element['attributes']))
            {
                foreach ($Element['attributes'] as $name => $value)
                {
                    if ($value === null)
                    {
                        continue;
                    }

                    $markup .= " $name=\"".self::escape($value).'"';
                }
            }
            */
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
            /*
            $markup .= $hasName ? '>' : '';
            */

            if (isset($Element['elements']))
            {
                $markup .= $this->elementsDocx($Element['elements']);
            }
            elseif (isset($Element['element']))
            {
                $markup .= $this->elementDocx($Element['element']);
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
                        if (isset($value['markdown'], $value['text']) && $value['markdown'] && $value['text'])
                        {
                            // value is markdown text, so parse it and handle elements
                            $tree = $this->getTree($value['text']);
                            if (isset($tree[0]['handler']['argument']))
                            {
                                $markup .= $this->lineDocx($tree[0]['handler']['argument']);
                            }
                        }
                    }
                    else
                    {
                        // value is text
                        if ($value)
                        {
                            $this->textRun->addText($value, $this->options);
                        }
                    }
                }
                else
                {
                    if ($text)
                    {
                        $this->textRun->addText(!$permitRawHtml ? self::escape($text, true) : $text, $this->options);
                    }
                }
            }

            if ($hasName)
            {
                if ($Element['name'] == 'strong')
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
                elseif ($Element['name'] == 'ul' || $Element['name'] == 'ol')
                {
                    $this->listStyleName = $parentListStyle;
                    $this->listDepth -= 1;
                }
            }

            /*
            $markup .= $hasName ? '</' . $Element['name'] . '>' : '';
            */
        }
        elseif ($hasName)
        {
            //$this->options = [];
            /*
            $markup .= ' />';
            */
        }

        return $markup;
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

        $markup = '';

        $parentList = null;
        $parentListStyle = null;

        if ($hasName)
        {
            if ($Element['name'] == 'p') {
                $this->p = new Paragraph();
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
                $this->options['underline'] = true;
            }
            elseif ($Element['name'] == 'img')
            {
                if (isset($Element['attributes']['src']))
                {
                    $height = 0;
                    $width = 0;
                    $options = [];
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

                    $p = new Paragraph();
                    $p->addImage($Element['attributes']['src'], "${width}cm", "${height}cm");
                }
            }
            elseif (in_array($Element['name'], ['h1', 'h2', 'h3', 'h4']))
            {
                $p = new Paragraph($this->getOdtParagraphStyle($Element['name']));
                $p->addText($Element['htext'], $this->getOdtTextStyle('bold'));
            }
            elseif ($Element['name'] == 'ol')
            {
                var_dump('>ul'.($this->listDepth+1));

                $parentListStyle = $this->listStyleName;
                $this->listDepth += 1;
                $this->listStyleName = 'roman';

                if ($this->listDepth > 0)
                {
                    $this->odtFinishedListItems[($this->listDepth-1).'_'.$this->list->getNumberOfElements()] = true;
                    //var_dump('After set:',$this->odtFinishedListItems);
                    $this->list->addParagraph($this->p);

                    $parentList = $this->list;
                }

                $this->list = new ODTList();
                if ($this->listDepth == 0)
                {
                    $this->list->setStyle($this->getOdtListStyle($this->listStyleName));
                }
            }
            elseif ($Element['name'] == 'ul')
            {
                var_dump('>ul'.($this->listDepth+1));

                $parentListStyle = $this->listStyleName;
                $this->listDepth += 1;
                $this->listStyleName = 'bullet';

                if ($this->listDepth > 0)
                {
                    $this->odtFinishedListItems[($this->listDepth-1).'_'.$this->list->getNumberOfElements()] = true;
                    //var_dump('After set:',$this->odtFinishedListItems);
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
                var_dump('>li'.$this->listDepth);
                $this->p = new Paragraph();
            }

            /*
            $markup .= '<' . $Element['name'];

            if (isset($Element['attributes']))
            {
                foreach ($Element['attributes'] as $name => $value)
                {
                    if ($value === null)
                    {
                        continue;
                    }

                    $markup .= " $name=\"".self::escape($value).'"';
                }
            }
            */
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
            /*
            $markup .= $hasName ? '>' : '';
            */

            if (isset($Element['elements']))
            {
                $markup .= $this->elementsOdt($Element['elements']);
            }
            elseif (isset($Element['element']))
            {
                $markup .= $this->elementOdt($Element['element']);
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
                        if (isset($value['markdown'], $value['text']) && $value['markdown'] && $value['text'])
                        {
                            // value is markdown text, so parse it and handle elements
                            $tree = $this->getTree($value['text']);
                            if (isset($tree[0]['handler']['argument']))
                            {
                                $markup .= $this->lineOdt($tree[0]['handler']['argument']);
                            }
                        }
                    }
                    else
                    {
                        // value is text
                        if ($value)
                        {
                            $this->p->addText($value, $this->wordOptionsToOdtTextStyle($this->options));
                        }
                    }
                }
                else
                {
                    if ($text)
                    {
                        $this->p->addText(!$permitRawHtml ? self::escape($text, true) : $text, $this->wordOptionsToOdtTextStyle($this->options));
                    }
                }
            }

            if ($hasName)
            {
                if ($Element['name'] == 'strong')
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
                elseif ($Element['name'] == 'ul' || $Element['name'] == 'ol')
                {
                    var_dump('<ul'.$this->listDepth);
                    if ($parentList && $this->listDepth > 0)
                    {
                        $parentList->addSubList($this->list);
                        $this->list = $parentList;
                    }
                    $this->listStyleName = $parentListStyle;
                    $this->listDepth -= 1;
                }
                elseif ($Element['name'] == 'li')
                {
                    var_dump('<li');

                    $key = $this->listDepth.'_'.($this->list->getNumberOfElements()-1);

                    if (!isset($this->odtFinishedListItems[$key]))
                    {
                        $this->list->addParagraph($this->p);
                    }
                    else
                    {
                        unset($this->odtFinishedListItems[$key]);
                        //var_dump('After unset:',$this->odtFinishedListItems);
                    }
                }
            }

            /*
            $markup .= $hasName ? '</' . $Element['name'] . '>' : '';
            */
        }
        elseif ($hasName)
        {
            //$this->options = [];
            /*
            $markup .= ' />';
            */
        }
        

        return $markup;
    }

    protected function elements(array $Elements)
    {
        $markup = '';

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

            $markup .= ($autoBreak ? "\n" : '') . $this->element($Element);
            $autoBreak = $autoBreakNext;
        }

        $markup .= $autoBreak ? "\n" : '';

        // TODO: REMOVE
        
        /*
        if ($markup)
        {
            var_dump('elements:',$Elements, $markup);
        }
        */

        return $markup;
    }

    protected function elementsDocx(array $Elements)
    {
        $markup = '';

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

            $markup .= ($autoBreak ? "\n" : '') . $this->elementDocx($Element);
            $autoBreak = $autoBreakNext;
        }

        $markup .= $autoBreak ? "\n" : '';

        return $markup;
    }
    
    protected function elementsOdt(array $Elements)
    {
        $markup = '';

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

            $markup .= ($autoBreak ? "\n" : '') . $this->elementOdt($Element);
            $autoBreak = $autoBreakNext;
        }

        $markup .= $autoBreak ? "\n" : '';

        return $markup;
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
    # Converts the options array used in PhpWord to an ODT text style object
    #

    protected function wordOptionsToOdtTextStyle(array $options)
    {
        $isBold = isset($options['bold']) && $options['bold'];
        $isItalic = isset($options['italic']) && $options['italic'];
        $isUnderline = isset($options['underline']) && $options['underline'] == 'single';

        if ($isBold && $isItalic && $isUnderline)
        {
            return isset($this->odtTextStyles['bold_italic_underline']) ? $this->odtTextStyles['bold_italic_underline'] : null;
        }

        if ($isBold && $isItalic)
        {
            return isset($this->odtTextStyles['bold_italic']) ? $this->odtTextStyles['bold_italic'] : null;
        }

        if ($isBold && $isUnderline)
        {
            return isset($this->odtTextStyles['bold_underline']) ? $this->odtTextStyles['bold_underline'] : null;
        }

        if ($isItalic && $isUnderline)
        {
            return isset($this->odtTextStyles['italic_underline']) ? $this->odtTextStyles['italic_underline'] : null;
        }

        if ($isBold)
        {
            return isset($this->odtTextStyles['bold']) ? $this->odtTextStyles['bold'] : null;
        }

        if ($isItalic)
        {
            return isset($this->odtTextStyles['italic']) ? $this->odtTextStyles['italic'] : null;
        }

        if ($isUnderline)
        {
            return isset($this->odtTextStyles['underline']) ? $this->odtTextStyles['underline'] : null;
        }

        return null;
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
    # Deprecated Methods
    #

    function parse($text)
    {
        $markup = $this->text($text);

        return $markup;
    }

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

    protected $VariableRegex = '/^\$\{((?:\\\\\$|[^$])+?)\}/us';

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
}
