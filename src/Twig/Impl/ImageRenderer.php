<?php

namespace PhpDocxTemplate\Twig\Impl;

use PhpDocxTemplate\PhpDocxTemplate;
use PhpDocxTemplate\Twig\RendererInterface;

class ImageRenderer implements RendererInterface
{
    private $template;

    public function __construct(PhpDocxTemplate $template)
    {
        $this->template = $template;
    }

    public function render(string $path, int $width, int $height, string $unit): string
    {
        $docx = $this->template->getDocx();
        $preparedImageAttrs = $docx->prepareImageAttrs(['path' => $path, 'width' => $width, 'height' => $height, 'unit' => $unit]);
        $imgPath = $preparedImageAttrs['src'];

        // get image index
        $imgIndex = $docx->getNextRelationsIndex($docx->getMainPartName());
        $rid = 'rId' . $imgIndex;

        // replace preparations
        $docx->addImageToRelations($docx->getMainPartName(), $rid, $imgPath, $preparedImageAttrs['mime']);
        $xmlImage = str_replace(array('{IMAGEID}', '{WIDTH}', '{HEIGHT}'), array($imgIndex, $preparedImageAttrs['width'], $preparedImageAttrs['height']), $docx->getImageTemplate());
        return $xmlImage;
    }
}
