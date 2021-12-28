<?php

namespace PhpDocxTemplate\Twig;

interface RenderListenerInterface
{
    public function notify(string $path, int $width, int $height): void;
}
