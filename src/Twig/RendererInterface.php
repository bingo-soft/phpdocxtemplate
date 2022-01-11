<?php

namespace PhpDocxTemplate\Twig;

interface RendererInterface
{
    public function render(string $path, int $width, int $height, string $unit): string;
}
