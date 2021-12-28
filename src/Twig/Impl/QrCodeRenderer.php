<?php

namespace PhpDocxTemplate\Twig\Impl;

use Endroid\QrCode\QrCode;
use Endroid\QrCode\Writer\PngWriter;
use PhpDocxTemplate\PhpDocxTemplate;
use PhpDocxTemplate\Twig\RendererInterface;

class QrCodeRenderer implements RendererInterface
{
    private $template;

    public function __construct(PhpDocxTemplate $template)
    {
        $this->template = $template;
    }

    public function render(string $url, int $size = 100, int $margin = 10): string
    {
        $qrCode = new QrCode($url);
        $qrCode->setSize($size);
        $qrCode->setMargin($margin);
        $writer = new PngWriter();
        $result = $writer->write($qrCode);
        $path = sys_get_temp_dir() . "/" . uniqid("", true) . ".png";
        $result->saveToFile($path);

        $docx = $this->template->getDocx();
        $preparedImageAttrs = $docx->prepareImageAttrs(['path' => $path, 'width' => $size, 'height' => $size]);
        $imgPath = $preparedImageAttrs['src'];

        // get image index
        $imgIndex = $docx->getNextRelationsIndex($docx->getMainPartName());
        $rid = 'rId' . $imgIndex;

        // replace preparations
        $docx->addImageToRelations($docx->getMainPartName(), $rid, $imgPath, $preparedImageAttrs['mime']);
        $xmlImage = str_replace(array('{RID}', '{WIDTH}', '{HEIGHT}'), array($rid, $preparedImageAttrs['width'], $preparedImageAttrs['height']), $docx->getImageTemplate());
        return $xmlImage;
    }
}
