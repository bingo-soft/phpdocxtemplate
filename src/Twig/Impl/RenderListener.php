<?php

namespace PhpDocxTemplate\Twig\Impl;

use PhpDocxTemplate\Twig\RenderListenerInterface;

class RenderListener implements RenderListenerInterface
{
    public function notify(string $path): void
    {
        echo $path, "\r\n";
    }
}
