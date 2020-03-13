<?php

namespace NS\TableExportBundle\Exporter;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Sheet
{
    /** @var string|null */
    private $title;

    /** @var string */
    private $template;

    /** @var array */
    private $parameters;

    public function __construct(string $template, ?array $parameters = null)
    {
        $this->template = $template;
        $this->parameters = $parameters ?? [];
    }

    public function setTitle(?string $title): void
    {
        $this->title = $title;
    }

    public function getTitle(): ?string
    {
        return $this->title;
    }

    /** @var string|null */
    private $safeTitle;

    public function getSafeTitle(): ?string
    {
        if (!$this->title) {
            return null;
        }

        if ($this->safeTitle) {
            return $this->safeTitle;
        }

        $name = str_replace(Worksheet::getInvalidCharacters(), ' ', $this->title);
        $this->safeTitle = strlen($name) >= 31 ? substr($name, 0, 25) . '...' : $this->title;

        return $this->safeTitle;
    }

    public function getTemplate(): string
    {
        return $this->template;
    }

    public function getParameters(): array
    {
        return $this->parameters;
    }
}
