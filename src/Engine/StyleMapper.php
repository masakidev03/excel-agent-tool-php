<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;

final class StyleMapper
{
    private const WRITE_PATTERNS = [
        'none' => 'none', 'solid' => 'solid', 'medium_gray' => 'mediumGray',
        'dark_gray' => 'darkGray', 'light_gray' => 'lightGray',
        'dark_horizontal' => 'darkHorizontal', 'dark_vertical' => 'darkVertical',
        'dark_down' => 'darkDown', 'dark_up' => 'darkUp', 'dark_grid' => 'darkGrid',
        'dark_trellis' => 'darkTrellis', 'light_horizontal' => 'lightHorizontal',
        'light_vertical' => 'lightVertical', 'light_down' => 'lightDown',
        'light_up' => 'lightUp', 'light_grid' => 'lightGrid', 'light_trellis' => 'lightTrellis',
        'gray_125' => 'gray125', 'gray_0625' => 'gray0625',
    ];

    /** @param array<string, mixed> $patch */
    public static function apply(Cell $cell, array $patch): void
    {
        $style = $cell->getStyle();
        if (isset($patch['number_format'])) {
            $style->getNumberFormat()->setFormatCode($patch['number_format']);
        }
        if (isset($patch['font'])) {
            $font = $style->getFont();
            $input = $patch['font'];
            if (isset($input['name'])) {
                $font->setName($input['name']);
            }
            if (isset($input['size'])) {
                $font->setSize((float) $input['size']);
            }
            if (array_key_exists('bold', $input)) {
                $font->setBold($input['bold']);
            }
            if (array_key_exists('italic', $input)) {
                $font->setItalic($input['italic']);
            }
            if (array_key_exists('underline', $input)) {
                $font->setUnderline($input['underline'] ? Font::UNDERLINE_SINGLE : Font::UNDERLINE_NONE);
            }
            if (isset($input['color'])) {
                $font->getColor()->setARGB(strtoupper($input['color']));
            }
        }
        if (isset($patch['fill'])) {
            $fill = $style->getFill();
            $input = $patch['fill'];
            if (isset($input['pattern'])) {
                $fill->setFillType(self::WRITE_PATTERNS[$input['pattern']]);
            }
            if (isset($input['foreground_color'])) {
                $fill->getStartColor()->setARGB(strtoupper($input['foreground_color']));
            }
            if (isset($input['background_color'])) {
                $fill->getEndColor()->setARGB(strtoupper($input['background_color']));
            }
        }
        if (isset($patch['alignment'])) {
            $alignment = $style->getAlignment();
            $input = $patch['alignment'];
            if (isset($input['horizontal'])) {
                $alignment->setHorizontal($input['horizontal']);
            }
            if (isset($input['vertical'])) {
                $alignment->setVertical($input['vertical']);
            }
            if (array_key_exists('wrap_text', $input)) {
                $alignment->setWrapText($input['wrap_text']);
            }
            if (isset($input['text_rotation'])) {
                $alignment->setTextRotation($input['text_rotation']);
            }
        }
    }

    /** @return array<string, mixed>|null */
    public static function observe(Cell $cell): ?array
    {
        if ($cell->getXfIndex() === 0) {
            return null;
        }
        $style = $cell->getStyle();
        $font = $style->getFont();
        $fill = $style->getFill();
        $alignment = $style->getAlignment();
        $result = [
            'style_id' => $cell->getXfIndex(),
            'number_format' => $style->getNumberFormat()->getFormatCode(),
            'font' => [
                'name' => $font->getName(),
                'size' => (float) $font->getSize(),
                'bold' => $font->getBold(),
                'italic' => $font->getItalic(),
                'underline' => $font->getUnderline() !== Font::UNDERLINE_NONE,
            ],
            'fill' => ['pattern' => self::observePattern($fill->getFillType())],
        ];
        $fontColor = $font->getColor()->getARGB();
        if (is_string($fontColor) && strlen($fontColor) === 8) {
            $result['font']['color'] = strtoupper($fontColor);
        }
        $startColor = $fill->getStartColor()->getARGB();
        if (is_string($startColor) && strlen($startColor) === 8 && $fill->getFillType() !== Fill::FILL_NONE) {
            $result['fill']['foreground_color'] = strtoupper($startColor);
        }
        $endColor = $fill->getEndColor()->getARGB();
        if (is_string($endColor) && strlen($endColor) === 8 && $fill->getFillType() !== Fill::FILL_NONE) {
            $result['fill']['background_color'] = strtoupper($endColor);
        }
        $observedAlignment = [];
        foreach (['horizontal' => $alignment->getHorizontal(), 'vertical' => $alignment->getVertical()] as $key => $value) {
            if ($value !== null) {
                $observedAlignment[$key] = $value;
            }
        }
        if ($alignment->getWrapText() || $alignment->getTextRotation() !== 0) {
            $observedAlignment['wrap_text'] = $alignment->getWrapText();
            $observedAlignment['text_rotation'] = $alignment->getTextRotation();
        }
        if ($observedAlignment !== []) {
            $result['alignment'] = $observedAlignment;
        }

        return $result;
    }

    private static function observePattern(string $value): string
    {
        $key = array_search($value, self::WRITE_PATTERNS, true);

        return $key === false ? $value : $key;
    }
}

