<?php

namespace PhpDocxTemplate\Escaper;

abstract class AbstractEscaper implements EscaperInterface
{
    /**
     * @param string $input
     *
     * @return string
     */
    abstract protected function escapeSingleValue(string $input): string;

    /**
     * @param mixed $input
     *
     * @return mixed
     */
    public function escape($input)
    {
        if (is_array($input)) {
            foreach ($input as &$item) {
                $item = $this->escapeSingleValue($item);
            }
        } else {
            $input = $this->escapeSingleValue($input);
        }

        return $input;
    }
}
