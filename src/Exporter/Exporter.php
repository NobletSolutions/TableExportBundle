<?php

namespace NS\TableExportBundle\Exporter;

use PhpOffice\PhpSpreadsheet\Reader\Html;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\StreamedResponse;
use Twig\Environment;

class Exporter
{
    private Environment $twig;

    public function __construct(Environment $twig)
    {
        $this->twig = $twig;
    }

    public function export(string $baseFileName, Sheet $sheet): StreamedResponse
    {
        $html = $this->twig->render($sheet->getTemplate(), $sheet->getParameters());
        $headers = [
            'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Pragma' => 'public',
            'Cache-Control' => 'maxage=1',
            'Content-Disposition' => sprintf('attachment;filename=%s-%s.xlsx', $baseFileName, date('Y_m_d_H_i_s')),
        ];

        return new StreamedResponse(static function () use ($html, $sheet) {
            $reader = new Html();
            $spreadsheet = $reader->loadFromString($html);
            if ($sheet->getSafeTitle()) {
                $spreadsheet->getActiveSheet()->setTitle($sheet->getSafeTitle());
            }

            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        }, Response::HTTP_OK, $headers);
    }

    /**
     * @param Sheet[]|array $sheets
     * @param string $baseFileName
     * @param array $headers
     * @return StreamedResponse
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     * @throws \Twig\Error\LoaderError
     * @throws \Twig\Error\RuntimeError
     * @throws \Twig\Error\SyntaxError
     */
    public function multiSheetExport(array $sheets, string $baseFileName, array $headers = []): StreamedResponse
    {
        $reader = new Html();
        $spreadsheet = null;
        $x = 0;

        foreach ($sheets as $sheet) {
            if ($spreadsheet) {
                $x++;
                $reader->setSheetIndex($x);
            }

            $spreadsheet = $reader->loadFromString($this->twig->render($sheet->getTemplate(), $sheet->getParameters()), $spreadsheet);
            if ($sheet->getSafeTitle()) {
                $spreadsheet->getActiveSheet()->setTitle($sheet->getSafeTitle());
            }
        }

        $spreadsheet->setActiveSheetIndex(0);

        $headers = array_merge([
            'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Pragma' => 'public',
            'Cache-Control' => 'maxage=1',
            'Content-Disposition' => sprintf('attachment;filename=%s-%s.xlsx', $baseFileName, date('Y_m_d_H_i_s')),
        ], $headers);

        return new StreamedResponse(static function () use ($spreadsheet) {
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        }, Response::HTTP_OK, $headers);
    }
}
