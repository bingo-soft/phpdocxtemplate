<?php

namespace PhpDocxTemplate\Escaper;

class RegExp extends AbstractEscaper
{
    private const REG_EXP_DELIMITER = '/';

    protected function escapeSingleValue(string $input): string
    {
        return self::REG_EXP_DELIMITER . preg_quote($input, self::REG_EXP_DELIMITER) . self::REG_EXP_DELIMITER . 'u';
    }
}
