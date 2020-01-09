[![Latest Stable Version](https://poser.pugx.org/bingo-soft/phpdocxtemplate/v/stable.png)](https://packagist.org/packages/bingo-soft/phpdocxtemplate)
[![Build Status](https://travis-ci.org/bingo-soft/phpdocxtemplate.png?branch=master)](https://travis-ci.org/bingo-soft/phpdocxtemplate)
[![Minimum PHP Version](https://img.shields.io/badge/php-%3E%3D%207.2-8892BF.svg)](https://php.net/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/bingo-soft/phpdocxtemplate/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/bingo-soft/phpdocxtemplate/?branch=master)
[![Code Coverage](https://scrutinizer-ci.com/g/bingo-soft/phpdocxtemplate/badges/coverage.png?b=master)](https://scrutinizer-ci.com/g/bingo-soft/phpdocxtemplate/?branch=master)

# PhpDocxTemplate

PhpDocxTemplate is a PHP library, which uses docx files as Twig templates

# Installation

Install PhpDocxTemplate, using Composer:

```
composer require bingo-soft/phpdocxtemplate
```

# Basic example

```php
use PhpDocxTemplate\PhpDocxTemplate;

$doc = "./templates/template1.docx";
$template = new PhpDocxTemplate($doc);
$template->render(["one" => "1", "two" => "2", "three" => "3", "four" => "4"]);
$template->save("./documents/report.docx");
```

## Acknowledgements

PhpDocxTemplate draws inspiration from the [python-docx-template](https://github.com/elapouya/python-docx-template) library.

## License

MIT
