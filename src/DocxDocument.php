<?php

namespace PhpDocxTemplate;

use DOMDocument;
use DOMElement;
use Exception;
use PhpZip\ZipFile;
use RecursiveIteratorIterator;
use RecursiveDirectoryIterator;

/**
 * Class DocxDocument
 *
 * @package PhpDocxTemplate
 */
class DocxDocument
{
    private $path;
    private $tmpDir = "./tmp/";
    private $document;

    /**
     * Construct an instance of Document
     *
     * @param string $path - path to the document
     * @param string $tmpPath - path to the temp directory
     *
     * @throws Exception
     */
    public function __construct(string $path, ?string $tmpPath = null)
    {
        if (!empty($tmpPath)) {
            if (!is_dir($tmpPath) && !mkdir($tmpPath, 0777, true)) {
                throw new Exception(
                    "The specified path \"" . $tmpPath . "\" for \"temp\" folder is not valid"
                );
            }

            $this->tmpDir = $tmpPath . "/";
        }

        if (file_exists($path)) {
            $this->path = $path;
            $this->tmpDir .= uniqid("", true) . date("His");
            $this->extract();
        } else {
            throw new Exception("The template " . $path . " was not found!");
        }
    }

    /**
     * Extract (unzip) document contents
     *
     * @throws \PhpZip\Exception\ZipException
     */
    private function extract(): void
    {
        if (file_exists($this->tmpDir) && is_dir($this->tmpDir)) {
            $this->rrmdir($this->tmpDir);
        }

        mkdir($this->tmpDir);

        $zip = new ZipFile();
        $zip->openFile($this->path);
        $zip->extractTo($this->tmpDir);
        $zip->close();

        $this->document = file_get_contents($this->tmpDir . "/word/document.xml");
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
     *
     * @throws \PhpZip\Exception\ZipException
     */
    public function save(string $path): void
    {
        $rootPath = realpath($this->tmpDir);

        $zip = new ZipFile();

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

        $zip->saveAsFile($path);
        $zip->close();

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
        $this->rrmdir($this->tmpDir);
    }
}
