<?php

namespace PhpDocxTemplate\Twig\Impl;

use Twig\Extension\AbstractExtension;
use Twig\Markup;
use Twig\TwigFunction;
use PhpDocxTemplate\Twig\RendererInterface;

class ImageExtension extends AbstractExtension
{
    private $renderer;

    /**
     * @return mixed
     */
    public function getFunctions()
    {
        return [
            new TwigFunction('image', [$this, 'image'])
        ];
    }

    public function setRenderer(RendererInterface $renderer): void
    {
        $this->renderer = $renderer;
    }

    public function image(string $path, ?int $width = 100, ?int $height = 100, ?string $unit = 'px'): object
    {
        return new Markup($this->renderer->render($path, $width, $height, $unit), 'UTF-8');
    }
}
