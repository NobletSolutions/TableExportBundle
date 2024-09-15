<?php

namespace NS\TableExportBundle\Exporter;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Sheet
{
    private string|null $title;

    private string $template;

    private array $parameters;

    public function __construct(string $template, ?array $parameters = null)
    {
        $this->template   = $template;
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

    private ?string $safeTitle = null;

    public function getSafeTitle(): ?string
    {
        if (!$this->title) {
            return null;
        }

        if ($this->safeTitle) {
            return $this->safeTitle;
        }

        $name            = str_replace(Worksheet::getInvalidCharacters(), ' ', $this->title);
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
