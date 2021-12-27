<?php

namespace PhpDocxTemplate;

use DOMDocument;
use DOMElement;
use Exception;
use ZipArchive;
use RecursiveIteratorIterator;
use RecursiveDirectoryIterator;
use PhpDocxTemplate\Escaper\RegExp;

/**
 * Class DocxDocument
 *
 * @package PhpDocxTemplate
 */
class DocxDocument
{
    private const MAXIMUM_REPLACEMENTS_DEFAULT = -1;
    private $path;
    private $tmpDir;
    private $document;
    private $zipClass;
    private $tempDocumentMainPart;
    private $tempDocumentHeaders = [];
    private $tempDocumentFooters = [];
    private $tempDocumentRelations = [];
    private $tempDocumentContentTypes = '';
    private $tempDocumentNewImages = [];

    /**
     * Construct an instance of Document
     *
     * @param string $path - path to the document
     *
     * @throws Exception
     */
    public function __construct(string $path)
    {
        if (file_exists($path)) {
            $this->path = $path;
            $this->tmpDir = sys_get_temp_dir() . "/" . uniqid("", true) . date("His");
            $this->zipClass = new ZipArchive();
            $this->extract();
        } else {
            throw new Exception("The template " . $path . " was not found!");
        }
    }

    /**
     * Extract (unzip) document contents
     */
    private function extract(): void
    {
        if (file_exists($this->tmpDir) && is_dir($this->tmpDir)) {
            $this->rrmdir($this->tmpDir);
        }

        mkdir($this->tmpDir);

        $this->zipClass->open($this->path);
        $this->zipClass->extractTo($this->tmpDir);

        $index = 1;
        while (false !== $this->zipClass->locateName($this->getHeaderName($index))) {
            $this->tempDocumentHeaders[$index] = $this->readPartWithRels($this->getHeaderName($index));
            $index += 1;
        }
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getFooterName($index))) {
            $this->tempDocumentFooters[$index] = $this->readPartWithRels($this->getFooterName($index));
            $index += 1;
        }

        $this->tempDocumentMainPart = $this->readPartWithRels($this->getMainPartName());

        $this->tempDocumentContentTypes = $this->zipClass->getFromName($this->getDocumentContentTypesName());

        //$this->zipClass->close();

        $this->document = file_get_contents($this->tmpDir . "/word/document.xml");
    }

    /**
     * Get document main part
     *
     * @return string
     */
    public function getDocumentMainPart(): string
    {
        return $this->tempDocumentMainPart;
    }

    /**
     * Get the name of main part document (method from PhpOffice\PhpWord)
     *
     * @return string
     */
    private function getMainPartName(): string
    {
        $contentTypes = $this->zipClass->getFromName('[Content_Types].xml');

        $pattern = '~PartName="\/(word\/document.*?\.xml)" ' .
                   'ContentType="application\/vnd\.openxmlformats-officedocument' .
                   '\.wordprocessingml\.document\.main\+xml"~';

        $matches = [];
        preg_match($pattern, $contentTypes, $matches);

        return array_key_exists(1, $matches) ? $matches[1] : 'word/document.xml';
    }

    /**
     * @return string
     */
    private function getDocumentContentTypesName(): string
    {
        return '[Content_Types].xml';
    }

    /**
     * Read document part (method from PhpOffice\PhpWord)
     *
     * @param string $fileName
     *
     * @return string
     */
    private function readPartWithRels(string $fileName): string
    {
        $relsFileName = $this->getRelationsName($fileName);
        $partRelations = $this->zipClass->getFromName($relsFileName);
        if ($partRelations !== false) {
            $this->tempDocumentRelations[$fileName] = $partRelations;
        }

        return $this->fixBrokenMacros($this->zipClass->getFromName($fileName));
    }

    /**
     * Get the name of the relations file for document part (method from PhpOffice\PhpWord)
     *
     * @param string $documentPartName
     *
     * @return string
     */
    private function getRelationsName(string $documentPartName): string
    {
        return 'word/_rels/' . pathinfo($documentPartName, PATHINFO_BASENAME) . '.rels';
    }

    private function getNextRelationsIndex(string $documentPartName): int
    {
        if (isset($this->tempDocumentRelations[$documentPartName])) {
            $candidate = substr_count($this->tempDocumentRelations[$documentPartName], '<Relationship');
            while (strpos($this->tempDocumentRelations[$documentPartName], 'Id="rId' . $candidate . '"') !== false) {
                $candidate++;
            }

            return $candidate;
        }

        return 1;
    }

    /**
     * Finds parts of broken macros and sticks them together (method from PhpOffice\PhpWord)
     *
     * @param string $documentPart
     *
     * @return string
     */
    private function fixBrokenMacros(string $documentPart): string
    {
        return preg_replace_callback(
            '/\$(?:\{|[^{$]*\>\{)[^}$]*\}/U',
            function ($match) {
                return strip_tags($match[0]);
            },
            $documentPart
        );
    }

    /**
     * @param string $macro
     *
     * @return string
     */
    protected static function ensureMacroCompleted(string $macro): string
    {
        if (substr($macro, 0, 2) !== '${' && substr($macro, -1) !== '}') {
            $macro = '${' . $macro . '}';
        }
        return $macro;
    }

    /**
     * Get the name of the header file for $index.
     *
     * @param int $index
     *
     * @return string
     */
    private function getHeaderName(int $index): string
    {
        return sprintf('word/header%d.xml', $index);
    }

    /**
     * Get the name of the footer file for $index.
     *
     * @param int $index
     *
     * @return string
     */
    private function getFooterName(int $index): string
    {
        return sprintf('word/footer%d.xml', $index);
    }

    /**
     * Find all variables in $documentPartXML.
     *
     * @param string $documentPartXML
     *
     * @return string[]
     */
    private function getVariablesForPart(string $documentPartXML): array
    {
        $matches = array();
        //preg_match_all('/\$\{(.*?)}/i', $documentPartXML, $matches);
        preg_match_all('/\{\{(.*?)\}\}/i', $documentPartXML, $matches);

        return $matches[1];
    }

    private function getImageArgs(string $varNameWithArgs): array
    {
        $varElements = explode(':', $varNameWithArgs);
        array_shift($varElements); // first element is name of variable => remove it

        $varInlineArgs = array();
        // size format documentation: https://msdn.microsoft.com/en-us/library/documentformat.openxml.vml.shape%28v=office.14%29.aspx?f=255&MSPPError=-2147217396
        foreach ($varElements as $argIdx => $varArg) {
            if (strpos($varArg, '=')) { // arg=value
                list($argName, $argValue) = explode('=', $varArg, 2);
                $argName = strtolower($argName);
                if ($argName == 'size') {
                    list($varInlineArgs['width'], $varInlineArgs['height']) = explode('x', $argValue, 2);
                } else {
                    $varInlineArgs[strtolower($argName)] = $argValue;
                }
            } elseif (preg_match('/^([0-9]*[a-z%]{0,2}|auto)x([0-9]*[a-z%]{0,2}|auto)$/i', $varArg)) { // 60x40
                list($varInlineArgs['width'], $varInlineArgs['height']) = explode('x', $varArg, 2);
            } else { // :60:40:f
                switch ($argIdx) {
                    case 0:
                        $varInlineArgs['width'] = $varArg;
                        break;
                    case 1:
                        $varInlineArgs['height'] = $varArg;
                        break;
                    case 2:
                        $varInlineArgs['ratio'] = $varArg;
                        break;
                }
            }
        }

        return $varInlineArgs;
    }

    /**
     * @param mixed $replaceImage
     * @param array $varInlineArgs
     *
     * @return array
     */
    private function prepareImageAttrs($replaceImage, array $varInlineArgs): array
    {
        // get image path and size
        $width = null;
        $height = null;
        $ratio = null;

        // a closure can be passed as replacement value which after resolving, can contain the replacement info for the image
        // use case: only when a image if found, the replacement tags can be generated
        if (is_callable($replaceImage)) {
            $replaceImage = $replaceImage();
        }

        if (is_array($replaceImage) && isset($replaceImage['path'])) {
            $imgPath = $replaceImage['path'];
            if (isset($replaceImage['width'])) {
                $width = $replaceImage['width'];
            }
            if (isset($replaceImage['height'])) {
                $height = $replaceImage['height'];
            }
            if (isset($replaceImage['ratio'])) {
                $ratio = $replaceImage['ratio'];
            }
        } else {
            $imgPath = $replaceImage;
        }

        $width = $this->chooseImageDimension($width, isset($varInlineArgs['width']) ? $varInlineArgs['width'] : null, 115);
        $height = $this->chooseImageDimension($height, isset($varInlineArgs['height']) ? $varInlineArgs['height'] : null, 70);

        $imageData = @getimagesize($imgPath);
        if (!is_array($imageData)) {
            throw new Exception(sprintf('Invalid image: %s', $imgPath));
        }
        list($actualWidth, $actualHeight, $imageType) = $imageData;

        // fix aspect ratio (by default)
        if (is_null($ratio) && isset($varInlineArgs['ratio'])) {
            $ratio = $varInlineArgs['ratio'];
        }
        if (is_null($ratio) || !in_array(strtolower($ratio), array('', '-', 'f', 'false'))) {
            $this->fixImageWidthHeightRatio($width, $height, $actualWidth, $actualHeight);
        }

        $imageAttrs = array(
            'src'    => $imgPath,
            'mime'   => image_type_to_mime_type($imageType),
            'width'  => $width,
            'height' => $height,
        );

        return $imageAttrs;
    }

    /**
     * @param mixed $width
     * @param mixed $height
     * @param int $actualWidth
     * @param int $actualHeight
     */
    private function fixImageWidthHeightRatio(&$width, &$height, int $actualWidth, int $actualHeight): void
    {
        $imageRatio = $actualWidth / $actualHeight;

        if (($width === '') && ($height === '')) { // defined size are empty
            $width = $actualWidth . 'px';
            $height = $actualHeight . 'px';
        } elseif ($width === '') { // defined width is empty
            $heightFloat = (float) $height;
            $widthFloat = $heightFloat * $imageRatio;
            $matches = array();
            preg_match("/\d([a-z%]+)$/", $height, $matches);
            $width = $widthFloat . $matches[1];
        } elseif ($height === '') { // defined height is empty
            $widthFloat = (float) $width;
            $heightFloat = $widthFloat / $imageRatio;
            $matches = array();
            preg_match("/\d([a-z%]+)$/", $width, $matches);
            $height = $heightFloat . $matches[1];
        } else { // we have defined size, but we need also check it aspect ratio
            $widthMatches = array();
            preg_match("/\d([a-z%]+)$/", $width, $widthMatches);
            $heightMatches = array();
            preg_match("/\d([a-z%]+)$/", $height, $heightMatches);
            // try to fix only if dimensions are same
            if ($widthMatches[1] == $heightMatches[1]) {
                $dimention = $widthMatches[1];
                $widthFloat = (float) $width;
                $heightFloat = (float) $height;
                $definedRatio = $widthFloat / $heightFloat;

                if ($imageRatio > $definedRatio) { // image wider than defined box
                    $height = ($widthFloat / $imageRatio) . $dimention;
                } elseif ($imageRatio < $definedRatio) { // image higher than defined box
                    $width = ($heightFloat * $imageRatio) . $dimention;
                }
            }
        }
    }

    private function chooseImageDimension(?int $baseValue, ?int $inlineValue, int $defaultValue): string
    {
        $value = $baseValue;
        if (is_null($value) && isset($inlineValue)) {
            $value = $inlineValue;
        }
        if (!preg_match('/^([0-9]*(cm|mm|in|pt|pc|px|%|em|ex|)|auto)$/i', $value)) {
            $value = null;
        }
        if (is_null($value)) {
            $value = $defaultValue;
        }
        if (is_numeric($value)) {
            $value .= 'px';
        }

        return $value;
    }

    private function addImageToRelations(string $partFileName, string $rid, string $imgPath, string $imageMimeType): void
    {
        // define templates
        $typeTpl = '<Override PartName="/word/media/{IMG}" ContentType="image/{EXT}"/>';
        $relationTpl = '<Relationship Id="{RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{IMG}"/>';
        $newRelationsTpl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n" . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        $newRelationsTypeTpl = '<Override PartName="/{RELS}" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $extTransform = array(
            'image/jpeg' => 'jpeg',
            'image/png'  => 'png',
            'image/bmp'  => 'bmp',
            'image/gif'  => 'gif',
        );

        // get image embed name
        if (isset($this->tempDocumentNewImages[$imgPath])) {
            $imgName = $this->tempDocumentNewImages[$imgPath];
        } else {
            // transform extension
            if (isset($extTransform[$imageMimeType])) {
                $imgExt = $extTransform[$imageMimeType];
            } else {
                throw new Exception("Unsupported image type $imageMimeType");
            }

            // add image to document
            $imgName = 'image_' . $rid . '_' . pathinfo($partFileName, PATHINFO_FILENAME) . '.' . $imgExt;
            $this->zipClass->addFile($imgPath, 'word/media/' . $imgName);

            $this->tempDocumentNewImages[$imgPath] = $imgName;

            // setup type for image
            $xmlImageType = str_replace(array('{IMG}', '{EXT}'), array($imgName, $imgExt), $typeTpl);
            $this->tempDocumentContentTypes = str_replace('</Types>', $xmlImageType, $this->tempDocumentContentTypes) . '</Types>';
        }

        $xmlImageRelation = str_replace(array('{RID}', '{IMG}'), array($rid, $imgName), $relationTpl);

        if (!isset($this->tempDocumentRelations[$partFileName])) {
            // create new relations file
            $this->tempDocumentRelations[$partFileName] = $newRelationsTpl;
            // and add it to content types
            $xmlRelationsType = str_replace('{RELS}', $this->getRelationsName($partFileName), $newRelationsTypeTpl);
            $this->tempDocumentContentTypes = str_replace('</Types>', $xmlRelationsType, $this->tempDocumentContentTypes) . '</Types>';
        }

        // add image to relations
        $this->tempDocumentRelations[$partFileName] = str_replace('</Relationships>', $xmlImageRelation, $this->tempDocumentRelations[$partFileName]) . '</Relationships>';
    }

    /**
     * @param mixed $search
     * @param mixed $replace Path to image, or array("path" => xx, "width" => yy, "height" => zz)
     * @param int $limit
     */
    public function setImageValue($search, $replace, ?int $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT): void
    {
        // prepare $search_replace
        if (!is_array($search)) {
            $search = array($search);
        }

        $replacesList = array();
        if (!is_array($replace) || isset($replace['path'])) {
            $replacesList[] = $replace;
        } else {
            $replacesList = array_values($replace);
        }

        $searchReplace = array();
        foreach ($search as $searchIdx => $searchString) {
            $searchReplace[$searchString] = isset($replacesList[$searchIdx]) ? $replacesList[$searchIdx] : $replacesList[0];
        }

        // collect document parts
        $searchParts = array(
            $this->getMainPartName() => &$this->tempDocumentMainPart,
        );
        foreach (array_keys($this->tempDocumentHeaders) as $headerIndex) {
            $searchParts[$this->getHeaderName($headerIndex)] = &$this->tempDocumentHeaders[$headerIndex];
        }
        foreach (array_keys($this->tempDocumentFooters) as $headerIndex) {
            $searchParts[$this->getFooterName($headerIndex)] = &$this->tempDocumentFooters[$headerIndex];
        }

        // define templates
        // result can be verified via "Open XML SDK 2.5 Productivity Tool" (http://www.microsoft.com/en-us/download/details.aspx?id=30425)
        $imgTpl = '<w:pict><v:shape type="#_x0000_t75" style="width:{WIDTH};height:{HEIGHT}" stroked="f"><v:imagedata r:id="{RID}" o:title=""/></v:shape></w:pict>';

        foreach ($searchParts as $partFileName => &$partContent) {
            $partVariables = $this->getVariablesForPart($partContent);

            foreach ($searchReplace as $searchString => $replaceImage) {
                $varsToReplace = array_filter($partVariables, function ($partVar) use ($searchString) {
                    return ($partVar == $searchString) || preg_match('/^' . preg_quote($searchString) . ':/', $partVar);
                });

                foreach ($varsToReplace as $varNameWithArgs) {
                    $varInlineArgs = $this->getImageArgs($varNameWithArgs);
                    $preparedImageAttrs = $this->prepareImageAttrs($replaceImage, $varInlineArgs);
                    $imgPath = $preparedImageAttrs['src'];

                    // get image index
                    $imgIndex = $this->getNextRelationsIndex($partFileName);
                    $rid = 'rId' . $imgIndex;

                    // replace preparations
                    $this->addImageToRelations($partFileName, $rid, $imgPath, $preparedImageAttrs['mime']);
                    $xmlImage = str_replace(array('{RID}', '{WIDTH}', '{HEIGHT}'), array($rid, $preparedImageAttrs['width'], $preparedImageAttrs['height']), $imgTpl);

                    // replace variable
                    $varNameWithArgsFixed = self::ensureMacroCompleted($varNameWithArgs);
                    $matches = array();
                    if (preg_match('/(<[^<]+>)([^<]*)(' . preg_quote($varNameWithArgsFixed) . ')([^>]*)(<[^>]+>)/Uu', $partContent, $matches)) {
                        $wholeTag = $matches[0];
                        array_shift($matches);
                        list($openTag, $prefix, , $postfix, $closeTag) = $matches;
                        $replaceXml = $openTag . $prefix . $closeTag . $xmlImage . $openTag . $postfix . $closeTag;
                        // replace on each iteration, because in one tag we can have 2+ inline variables => before proceed next variable we need to change $partContent
                        $partContent = $this->setValueForPart($wholeTag, $replaceXml, $partContent, $limit);
                    }
                }
            }
        }
    }

    /**
     * Find and replace macros in the given XML section.
     *
     * @param mixed $search
     * @param mixed $replace
     * @param string $documentPartXML
     * @param int $limit
     *
     * @return string
     */
    protected function setValueForPart($search, $replace, string $documentPartXML, int $limit): string
    {
        // Note: we can't use the same function for both cases here, because of performance considerations.
        if (self::MAXIMUM_REPLACEMENTS_DEFAULT === $limit) {
            return str_replace($search, $replace, $documentPartXML);
        }
        $regExpEscaper = new RegExp();

        return preg_replace($regExpEscaper->escape($search), $replace, $documentPartXML, $limit);
    }

    /**
     * Get document.xml contents as DOMDocument
     *
     * @return DOMDocument
     */
    public function getDOMDocument(): DOMDocument
    {
        $dom = new DOMDocument();
        $dom->loadXML($this->document);
        return $dom;
    }

    /**
     * Update document.xml contents
     *
     * @param DOMDocument $dom - new contents
     */
    public function updateDOMDocument(DOMDocument $dom): void
    {
        $this->document = $dom->saveXml();
        file_put_contents($this->tmpDir . "/word/document.xml", $this->document);
    }

    /**
     * Fix table corruption
     *
     * @param string $xml - xml to fix
     *
     * @return DOMDocument
     */
    public function fixTables(string $xml): DOMDocument
    {
        $dom = new DOMDocument();
        $dom->loadXML($xml);
        $tables = $dom->getElementsByTagName('tbl');
        foreach ($tables as $table) {
            $columns = [];
            $columnsLen = 0;
            $toAdd = 0;
            $tableGrid = null;
            foreach ($table->childNodes as $el) {
                if ($el->nodeName == 'w:tblGrid') {
                    $tableGrid = $el;
                    foreach ($el->childNodes as $col) {
                        if ($col->nodeName == 'w:gridCol') {
                            $columns[] = $col;
                            $columnsLen += 1;
                        }
                    }
                } elseif ($el->nodeName == 'w:tr') {
                    $cellsLen = 0;
                    foreach ($el->childNodes as $col) {
                        if ($col->nodeName == 'w:tc') {
                            $cellsLen += 1;
                        }
                    }
                    if (($columnsLen + $toAdd) < $cellsLen) {
                        $toAdd = $cellsLen - $columnsLen;
                    }
                }
            }

            // add columns, if necessary
            if (!is_null($tableGrid) && $toAdd > 0) {
                $width = 0;
                foreach ($columns as $col) {
                    if (!is_null($col->getAttribute('w:w'))) {
                        $width += $col->getAttribute('w:w');
                    }
                }
                if ($width > 0) {
                    $oldAverage = $width / $columnsLen;
                    $newAverage = round($width / ($columnsLen + $toAdd));
                    foreach ($columns as $col) {
                        $col->setAttribute('w:w', round($col->getAttribute('w:w') * $newAverage / $oldAverage));
                    }
                    while ($toAdd > 0) {
                        $newCol = $dom->createElement("w:gridCol");
                        $newCol->setAttribute('w:w', $newAverage);
                        $tableGrid->appendChild($newCol);
                        $toAdd -= 1;
                    }
                }
            }

            // remove columns, if necessary
            $columns = [];
            foreach ($tableGrid->childNodes as $col) {
                if ($col->nodeName == 'w:gridCol') {
                    $columns[] = $col;
                }
            }
            $columnsLen = count($columns);

            $cellsLen = 0;
            $cellsLenMax = 0;
            foreach ($table->childNodes as $el) {
                if ($el->nodeName == 'w:tr') {
                    $cells = [];
                    foreach ($el->childNodes as $col) {
                        if ($col->nodeName == 'w:tc') {
                            $cells[] = $col;
                        }
                    }
                    $cellsLen = $this->getCellLen($cells);
                    $cellsLenMax = max($cellsLenMax, $cellsLen);
                }
            }
            $toRemove = $cellsLen - $cellsLenMax;
            if ($toRemove > 0) {
                $removedWidth = 0.0;
                for ($i = $columnsLen - 1; ($i + 1) >= $toRemove; $i -= 1) {
                    $extraCol = $columns[$i];
                    $removedWidth += $extraCol->getAttribute('w:w');
                    $tableGrid->removeChild($extraCol);
                }

                $columnsLeft = [];
                foreach ($tableGrid->childNodes as $col) {
                    if ($col->nodeName == 'w:gridCol') {
                        $columnsLeft[] = $col;
                    }
                }
                $extraSpace = 0;
                if (count($columnsLeft) > 0) {
                    $extraSpace = $removedWidth / count($columnsLeft);
                }
                foreach ($columnsLeft as $col) {
                    $col->setAttribute('w:w', round($col->getAttribute('w:w') + $extraSpace));
                }
            }
        }
        return $dom;
    }

    /**
     * Get total cells length
     *
     * @param array $cells - cells
     *
     * @return int
     */
    private function getCellLen(array $cells): int
    {
        $total = 0;
        foreach ($cells as $cell) {
            foreach ($cell->childNodes as $tc) {
                if ($tc->nodeName == 'w:tcPr') {
                    foreach ($tc->childNodes as $span) {
                        if ($span->nodeName == 'w:gridSpan') {
                            $total += intval($span->getAttribute('w:val'));
                            break;
                        }
                    }
                    break;
                }
            }
        }
        return $total + 1;
    }

    /**
     * Save the document to the target path
     *
     * @param string $path - target path
     */
    public function save(string $path): void
    {
        $rootPath = realpath($this->tmpDir);

        $zip = new ZipArchive();
        $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($rootPath),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach ($files as $name => $file) {
            if (!$file->isDir()) {
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($rootPath) + 1);
                $zip->addFile($filePath, $relativePath);
            }
        }

        $zip->close();

        if (isset($this->zipClass)) {
            $this->zipClass->close();
        }

        $this->rrmdir($this->tmpDir);
    }

    /**
     * Remove recursively directory
     *
     * @param string $dir - target directory
     */
    private function rrmdir(string $dir): void
    {
        $objects = scandir($dir);
        if (is_array($objects)) {
            foreach ($objects as $object) {
                if ($object != "." && $object != "..") {
                    if (filetype($dir . "/" . $object) == "dir") {
                        $this->rrmdir($dir . "/" . $object);
                    } else {
                        unlink($dir . "/" . $object);
                    }
                }
            }
            reset($objects);
            rmdir($dir);
        }
    }

    /**
     * Close document
     */
    public function close(): void
    {
        if (isset($this->zipClass)) {
            $this->zipClass->close();
        }
        $this->rrmdir($this->tmpDir);
    }
}
