<?php

namespace PhpDocxTemplate\Twig\Impl;

use Twig\Extension\AbstractExtension;
use Twig\TwigFunction;
use PhpDocxTemplate\Twig\RenderListenerInterface;

class ImageExtension extends AbstractExtension
{
    private $listeners = [];

    /**
     * @return mixed
     */
    public function getFunctions()
    {
        return [
            new TwigFunction('image', [$this, 'image'])
        ];
    }

    public function addListener(RenderListenerInterface $listener): void
    {
        $this->listeners[] = $listener;
    }

    public function image(string $path, ?int $width = 100, ?int $height = 100): string
    {
        foreach ($this->listeners as $listener) {
            $listener->notify($path, $width, $height);
        }
        return 111;
    }
}
