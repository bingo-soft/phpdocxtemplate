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

    public function render(string $path, int $width, int $height): string
    {
        $docx = $this->template->getDocx();
        $preparedImageAttrs = $docx->prepareImageAttrs(['path' => $path, 'width' => $width, 'height' => $height]);
        $imgPath = $preparedImageAttrs['src'];

        // get image index
        $imgIndex = $docx->getNextRelationsIndex($path);
        $rid = 'rId' . $imgIndex;

        // replace preparations
        $docx->addImageToRelations($path, $rid, $imgPath, $preparedImageAttrs['mime']);
        $xmlImage = str_replace(array('{RID}', '{WIDTH}', '{HEIGHT}'), array($rid, $preparedImageAttrs['width'], $preparedImageAttrs['height']), $docx->getImageTemplate());
        return $xmlImage;
    }
}
