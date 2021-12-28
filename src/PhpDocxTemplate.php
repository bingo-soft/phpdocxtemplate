<?php

namespace PhpDocxTemplate;

use DOMDocument;
use DOMElement;
use Twig\Loader\ArrayLoader;
use Twig\Environment;
use PhpDocxTemplate\Twig\Impl\{
    ImageExtension,
    RenderListener
};

/**
 * Class PhpDocxTemplate
 *
 * @package PhpDocxTemplate
 */
class PhpDocxTemplate
{
    private const NEWLINE_XML = '</w:t><w:br/><w:t xml:space="preserve">';
    private const NEWPARAGRAPH_XML = '</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">';
    private const TAB_XML = '</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">';
    private const PAGE_BREAK = '</w:t><w:br w:type="page"/><w:t xml:space="preserve">';

    private const HEADER_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    private const FOOTER_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

    private $docx;
    private $crcToNewMedia;
    private $crcToNewEmbedded;
    private $picToReplace;
    private $picMap;

    /**
     * Construct an instance of PhpDocxTemplate
     *
     * @param string $path - path to the template
     */
    public function __construct(string $path)
    {
        $this->docx = new DocxDocument($path);
        $this->crcToNewMedia = [];
        $this->crcToNewEmbedded = [];
        $this->picToReplace = [];
        $this->picMap = [];
    }

    /**
     * Convert DOM to string
     *
     * @param DOMDocument $dom - DOM to be converted
     *
     * @return string
     */
    public function xmlToString(DOMDocument $dom): string
    {
        //return $el->ownerDocument->saveXML($el);
        return $dom->saveXML();
    }

    /**
     * Get document wrapper
     *
     * @return DocxDocument
     */
    public function getDocx(): DocxDocument
    {
        return $this->docx;
    }

    /**
     * Convert document.xml contents as string
     *
     * @return string
     */
    public function getXml(): string
    {
        return $this->xmlToString($this->docx->getDOMDocument());
    }

    /**
     * Write document.xml contents to file
     */
    private function writeXml(string $path): void
    {
        file_put_contents($path, $this->getXml());
    }

    /**
     * Update document.xml contents to file
     *
     * @param DOMDocument $xml - new contents
     */
    private function updateXml(DOMDocument $xml): void
    {
        $this->docx->updateDOMDocument($xml);
    }

    /**
     * Patch initial xml
     *
     * @param string $xml - initial xml
     */
    public function patchXml(string $xml): string
    {
        $xml = preg_replace('/(?<={)(<[^>]*>)+(?=[\{%])|(?<=[%\}])(<[^>]*>)+(?=\})/mu', '', $xml);
        $xml = preg_replace_callback(
            '/{%(?:(?!%}).)*|{{(?:(?!}}).)*/mu',
            array(get_class($this), 'stripTags'),
            $xml
        );
        $xml = preg_replace_callback(
            '/(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*colspan\s+([^%]*)\s*%}(.*?<\/w:tc>)/mu',
            array(get_class($this), 'colspan'),
            $xml
        );
        $xml = preg_replace_callback(
            '/(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?<\/w:tc>)/mu',
            array(get_class($this), 'cellbg'),
            $xml
        );
        // avoid {{r and {%r tags to strip MS xml tags too far
        // ensure space preservation when splitting
        $xml = preg_replace(
            '/<w:t>((?:(?!<w:t>).)*)({{r\s.*?}}|{%r\s.*?%})/mu',
            '<w:t xml:space="preserve">${1}${2}',
            $xml
        );
        $xml = preg_replace(
            '/({{r\s.*?}}|{%r\s.*?%})/mu',
            '</w:t></w:r><w:r><w:t xml:space="preserve">${1}</w:t></w:r><w:r><w:t xml:space="preserve">',
            $xml
        );

        // {%- will merge with previous paragraph text
        $xml = preg_replace(
            '/<\/w:t>(?:(?!<\/w:t>).)*?{%-/mu',
            '{%',
            $xml
        );

        // -%} will merge with next paragraph text
        $xml = preg_replace(
            '/-%}(?:(?!<w:t[ >]).)*?<w:t[^>]*?>/mu',
            '%}',
            $xml
        );

        // replace into xml code the row/paragraph/run containing
        // {%y xxx %} or {{y xxx}} template tag
        // by {% xxx %} or {{ xx }} without any surronding <w:y> tags
        $tokens = ['tr', 'tc', 'p', 'r'];
        foreach ($tokens as $token) {
            $regex = '/';
            $regex .= str_replace("%s", $token, '<w:%s[ >](?:(?!<w:%s[ >]).)*({%|{{)%s ([^}%]*(?:%}|}})).*?<\/w:%s>');
            $regex .= '/mu';
            $xml = preg_replace(
                $regex,
                '${1} ${2}',
                $xml
            );
        }

        $xml = preg_replace_callback(
            '/<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*vm\s*%}.*?<\/w:tc[ >]/mu',
            array(get_class($this), 'vMergeTc'),
            $xml
        );

        $xml = preg_replace_callback(
            '/<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*hm\s*%}.*?<\/w:tc[ >]/mu',
            array(get_class($this), 'hMergeTc'),
            $xml
        );

        $xml = preg_replace_callback(
            '/(?<=\{[\{%])(.*?)(?=[\}%]})/mu',
            array(get_class($this), 'cleanTags'),
            $xml
        );

        return $xml;
    }

    private function resolveListing(string $xml): string
    {
        return preg_replace_callback(
            '/<w:p\b(?:[^>]*)?>.*?<\/w:p>/mus',
            array(get_class($this), 'resolveParagraph'),
            $xml
        );
    }

    private function resolveParagraph(array $matches): string
    {
        preg_match("/<w:pPr>.*<\/w:pPr>/mus", $matches[0], $paragraphProperties);

        return preg_replace_callback(
            '/<w:r\b(?:[^>]*)?>.*?<\/w:r>/mus',
            function ($m) use ($paragraphProperties) {
                return $this->resolveRun($paragraphProperties[0] ?? '', $m);
            },
            $matches[0]
        );
    }

    private function resolveRun(string $paragraphProperties, array $matches): string
    {
        preg_match("/<w:rPr>.*<\/w:rPr>/mus", $matches[0], $runProperties);

        return preg_replace_callback(
            '/<w:t\b(?:[^>]*)?>.*?<\/w:t>/mus',
            function ($m) use ($paragraphProperties, $runProperties) {
                return $this->resolveText($paragraphProperties, $runProperties[0] ?? '', $m);
            },
            $matches[0]
        );
    }

    private function resolveText(string $paragraphProperties, string $runProperties, array $matches): string
    {
        $xml = str_replace(
            "\t",
            sprintf("</w:t></w:r>" .
                "<w:r>%s<w:tab/></w:r>" .
                "<w:r>%s<w:t xml:space=\"preserve\">", $runProperties, $runProperties),
            $matches[0]
        );

        $xml = str_replace(
            "\a",
            sprintf("</w:t></w:r></w:p>" .
                "<w:p>%s<w:r>%s<w:t xml:space=\"preserve\">", $paragraphProperties, $runProperties),
            $xml
        );

        $xml = str_replace("\n", sprintf("</w:t>" .
            "</w:r>" .
            "</w:p>" .
            "<w:p>%s" .
            "<w:r>%s" .
            "<w:t xml:space=\"preserve\">", $paragraphProperties, $runProperties), $xml);

        $xml = str_replace(
            "\f",
            sprintf("</w:t></w:r></w:p>" .
                "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>" .
                "<w:p>%s<w:r>%s<w:t xml:space=\"preserve\">", $paragraphProperties, $runProperties),
            $xml
        );

        return $xml;
    }

    /**
     * Strip tags from matches
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function stripTags(array $matches): string
    {
        return preg_replace('/<\/w:t>.*?(<w:t>|<w:t [^>]*>)/mu', '', $matches[0]);
    }

    /**
     * Parse colspan
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function colspan(array $matches): string
    {
        $cellXml = $matches[1] . $matches[3];
        $cellXml = preg_replace('/<w:r[ >](?:(?!<w:r[ >]).)*<w:t><\/w:t>.*?<\/w:r>/mu', '', $cellXml);
        $cellXml = preg_replace('/<w:gridSpan[^\/]*\/>/mu', '', $cellXml, 1);
        return preg_replace(
            '/(<w:tcPr[^>]*>)/mu',
            sprintf('${1}<w:gridSpan w:val="{{%s}}"/>', $matches[2]),
            $cellXml
        );
    }

    /**
     * Parse cellbg
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function cellbg(array $matches): string
    {
        $cellXml = $matches[1] . $matches[3];
        $cellXml = preg_replace('/<w:r[ >](?:(?!<w:r[ >]).)*<w:t><\/w:t>.*?<\/w:r>/mu', '', $cellXml);
        $cellXml = preg_replace('/<w:shd[^\/]*\/>/mu', '', $cellXml, 1);
        return preg_replace(
            '/(<w:tcPr[^>]*>)/mu',
            sprintf('${1}<w:shd w:val="clear" w:color="auto" w:fill="{{%s}}"/>', $matches[2]),
            $cellXml
        );
    }

    /**
     * Parse vm
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function vMergeTc(array $matches): string
    {
        return preg_replace_callback(
            '/(<\/w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*vm\s*%})(.*?)(<\/w:t>)/mu',
            array(get_class($this), 'vMerge'),
            $matches[0]
        );
    }

    /**
     * Continue parsing vm
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function vMerge(array $matches): string
    {
        return '<w:vMerge w:val="{% if loop.first %}restart{% else %}continue{% endif %}"/>' .
            $matches[1] .  // Everything between ``</w:tcPr>`` and ``<w:t>``.
            "{% if loop.first %}" .
            $matches[2] .  // Everything before ``{% vm %}``.
            $matches[3] .  // Everything after ``{% vm %}``.
            "{% endif %}" .
            $matches[4];  // ``</w:t>``.
    }

    /**
     * Parse hm
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function hMergeTc(array $matches): string
    {
        $xmlToPatch = $matches[0];
        if (strpos($xmlToPatch, 'w:gridSpan') !== false) {
            $xmlToPatch = preg_replace_callback(
                '/(w:gridSpan w:val=")(\d+)(")/mu',
                array(get_class($this), 'withGridspan'),
                $xmlToPatch
            );
            $xmlToPatch = preg_replace('/{%\s*hm\s*%}/mu', '', $xmlToPatch);
        } else {
            $xmlToPatch = preg_replace_callback(
                '/(<\/w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*hm\s*%})(.*?)(<\/w:t>)/mu',
                array(get_class($this), 'withoutGridspan'),
                $xmlToPatch
            );
        }

        return "{% if loop.first %}" . $xmlToPatch . "{% endif %}";
    }

    private function withGridspan(array $matches): string
    {
        return $matches[1] . // ``w:gridSpan w:val="``.
            '{{ ' . $matches[2] . ' * loop.length }}' . // Content of ``w:val``, multiplied by loop length.
            $matches[3];  // Closing quotation mark.
    }

    private function withoutGridspan(array $matches): string
    {
        return '<w:gridSpan w:val="{{ loop.length }}"/>' .
            $matches[1] . // Everything between ``</w:tcPr>`` and ``<w:t>``.
            $matches[2] . // Everything before ``{% hm %}``.
            $matches[3] . // Everything after ``{% hm %}``.
            $matches[4]; // ``</w:t>``.
    }

    /**
     * Clean tags in matches
     *
     * @param array $matches - matches
     *
     * @return string
     */
    private function cleanTags(array $matches): string
    {
        return str_replace(
            ["&#8216;", '&lt;', '&gt;', '“', '”', "‘", "’"],
            ["'", '<', '>', '"', '"', "'", "'"],
            $matches[0]
        );
    }

    /**
     * Render xml
     *
     * @param string $srcXml - source xml
     * @param array $context - data to be rendered
     *
     * @return string
     */
    private function renderXml(string $srcXml, array $context): string
    {
        $srcXml = str_replace('<w:p>', "\n<w:p>", $srcXml);

        $ext = new ImageExtension();
        $ext->addListener(
            new RenderListener($this)
        );

        $template = new Environment(new ArrayLoader([
            'index' => $srcXml,
        ]));
        $template->addExtension($ext);

        $dstXml = $template->render('index', $context);

        $dstXml = str_replace(
            ["\n<w:p>", "{_{", '}_}', '{_%', '%_}'],
            ['<w:p>', "{{", '}}', '{%', '%}'],
            $dstXml
        );

        // fix xml after rendering
        $dstXml = preg_replace(
            '/<w:p [^>]*>(?:<w:r [^>]*><w:t [^>]*>\s*<\/w:t><\/w:r>)?(?:<w:pPr><w:ind w:left="360"\/>' .
            '<\/w:pPr>)?<w:r [^>]*>(?:<w:t\/>|<w:t [^>]*><\/w:t>|<w:t [^>]*\/>|<w:t><\/w:t>)<\/w:r><\/w:p>/mu',
            '',
            $dstXml
        );

        $dstXml = $this->resolveListing($dstXml);

        return $dstXml;
    }

    /**
     * Build xml
     *
     * @param array $context - data to be rendered
     *
     * @return string
     */
    public function buildXml(array $context): string
    {
        $xml = $this->getXml();
        $xml = $this->patchXml($xml);
        $xml = $this->renderXml($xml, $context);
        return $xml;
    }

    /**
     * Render document
     *
     * @param array $context - data to be rendered
     */
    public function render(array $context): void
    {
        $xmlSrc = $this->buildXml($context);
        $newXml = $this->docx->fixTables($xmlSrc);
        $this->updateXml($newXml);
    }

    /**
     * Save document
     *
     * @param string $path - target path
     */
    public function save(string $path): void
    {
        //$this->preProcessing();
        $this->docx->save($path);
        //$this->postProcessing($path);
    }

    /**
     * Clean everything after rendering
     */
    public function close(): void
    {
        $this->docx->close();
    }

    /**
     * @param mixed $search
     * @param mixed $replace Path to image, or array("path" => xx, "width" => yy, "height" => zz)
     * @param int $limit
     */
    public function setImageValue($search, $replace, ?int $limit = null): void
    {
        $this->docx->setImageValue($search, $replace, $limit = null);
    }
}
