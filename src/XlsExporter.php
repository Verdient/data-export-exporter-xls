<?php

declare(strict_types=1);

namespace Verdient\Hyperf3\DataExport\Exporter;

use Verdient\Hyperf3\DataExport\DataCollectorInterface;
use Verdient\Hyperf3\DataExport\DataExporterInterface;
use Verdient\Hyperf3\Logger\HasLogger;
use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

/**
 * 电子表格导出器
 * @author Verdient。
 */
class XlsExporter implements DataExporterInterface
{
    use HasLogger;

    /**
     * @inheritdoc
     * @author Verdient。
     */
    public function export(
        DataCollectorInterface $collector,
        ?string $path = null
    ): string|false {
        if ($path) {
            $dir = dirname($path);
            $filename = basename($path);
        } else {
            $dir = implode(DIRECTORY_SEPARATOR, [constant('BASE_PATH'), 'runtime', 'data', 'export']);
            $filename = $collector->fileName();
        }

        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }

        $excel  = new Excel([
            'path' => $dir
        ]);

        $constMemory = $collector->estimate() > 5000;
        if ($constMemory) {
            $this->logger()->info('使用固定内存模式');
            $fileObject = $excel->constMemory($filename, 'Sheet1', false);
            $fileHandle = $fileObject->getHandle();
            $format = new Format($fileHandle);
            $headerStyle = $format
                ->bold()
                ->align(Format::FORMAT_ALIGN_VERTICAL_CENTER)->toResource();
            $fileObject
                ->setType([Excel::TYPE_STRING])
                ->setRow('A1', 20, $headerStyle)
                ->header($collector->headers());
        } else {
            $fileObject = $excel->fileName($filename);
            $fileObject->setCurrentLine(1);
        }
        $count = 0;
        foreach ($collector->collect() as $row) {
            foreach ($row as $index2 => $value) {
                if (is_numeric($value)) {
                    if ($value > 99999999999) {
                        $row[$index2] = strval($value);
                    } else {
                        $value = strval($value);
                        if (str_contains($value, '.')) {
                            $row[$index2] = floatval($value);
                        } else {
                            $row[$index2] = intval($value);
                        }
                    }
                }
            }
            $fileObject->data([$row]);
            $count++;
            if ($count % 10000 === 0) {
                $this->logger()->info('已处理 ' . $count . ' 条');
            }
        }
        if (!$constMemory) {
            $fileHandle = $fileObject->getHandle();
            $format = new Format($fileHandle);
            $headerStyle = $format
                ->bold()
                ->align(Format::FORMAT_ALIGN_VERTICAL_CENTER)->toResource();
            $fileObject
                ->setType([Excel::TYPE_STRING])
                ->setRow('A1', 20, $headerStyle)
                ->header($collector->headers());
        }
        $this->logger()->info('本次导出完成，共计导出数据 ' . $count . ' 条');
        return $fileObject->output();
    }
}
