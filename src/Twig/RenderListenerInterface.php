<?php

namespace PhpDocxTemplate\Twig;

interface RenderListenerInterface
{
    public function notify(string $path): void;
}
