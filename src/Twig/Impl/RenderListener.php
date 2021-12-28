<?php

namespace PhpDocxTemplate\Twig\Impl;

use PhpDocxTemplate\PhpDocxTemplate;
use PhpDocxTemplate\Twig\RenderListenerInterface;

class RenderListener implements RenderListenerInterface
{
    private $template;

    public function __construct(PhpDocxTemplate $template)
    {
        $this->template = $template;
    }

    public function notify(string $path, int $width, int $height): void
    {
        var_dump($path);
    }
}
