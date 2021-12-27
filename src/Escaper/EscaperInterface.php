<?php

namespace PhpDocxTemplate\Escaper;

interface EscaperInterface
{
    /**
     * @param mixed $input
     *
     * @return mixed
     */
    public function escape($input);
}
