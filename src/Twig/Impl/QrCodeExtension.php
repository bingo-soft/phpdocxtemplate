<?php

namespace PhpDocxTemplate\Twig\Impl;

use Twig\Extension\AbstractExtension;
use Twig\Markup;
use Twig\TwigFunction;
use PhpDocxTemplate\Twig\RendererInterface;

class QrCodeExtension extends AbstractExtension
{
    private $renderer;

    /**
     * @return mixed
     */
    public function getFunctions()
    {
        return [
            new TwigFunction('qr_code', [$this, 'qrCode'])
        ];
    }

    public function setRenderer(RendererInterface $renderer): void
    {
        $this->renderer = $renderer;
    }

    public function qrCode(string $url, int $size = 100, int $margin = 10, string $unit = 'px'): object
    {
        return new Markup($this->renderer->render($url, $size, $margin, $unit), 'UTF-8');
    }
}
