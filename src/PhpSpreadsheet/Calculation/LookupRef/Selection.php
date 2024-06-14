<?php

namespace PhpOffice\PhpSpreadsheet\Calculation\LookupRef;

use PhpOffice\PhpSpreadsheet\Calculation\ArrayEnabled;
use PhpOffice\PhpSpreadsheet\Calculation\Functions;
use PhpOffice\PhpSpreadsheet\Calculation\Information\ExcelError;
use PhpOffice\PhpSpreadsheet\Calculation\LookupRef\RowColumnInformation;
use PhpOffice\PhpSpreadsheet\Calculation\LookupRef\Matrix;

class Selection
{
    use ArrayEnabled;

    /**
     * CHOOSE.
     *
     * Uses lookup_value to return a value from the list of value arguments.
     * Use CHOOSE to select one of up to 254 values based on the lookup_value.
     *
     * Excel Function:
     *        =CHOOSE(index_num, value1, [value2], ...)
     *
     * @param mixed $chosenEntry The entry to select from the list (indexed from 1)
     * @param mixed ...$chooseArgs Data values
     *
     * @return mixed The selected value
     */
    public static function choose(mixed $chosenEntry, mixed ...$chooseArgs): mixed
    {
        if (is_array($chosenEntry)) {
            return self::evaluateArrayArgumentsSubset([self::class, __FUNCTION__], 1, $chosenEntry, ...$chooseArgs);
        }

        $entryCount = count($chooseArgs) - 1;

        if (is_numeric($chosenEntry)) {
            --$chosenEntry;
        } else {
            return ExcelError::VALUE();
        }
        $chosenEntry = floor($chosenEntry);
        if (($chosenEntry < 0) || ($chosenEntry > $entryCount)) {
            return ExcelError::VALUE();
        }

        if (is_array($chooseArgs[$chosenEntry])) {
            return Functions::flattenArray($chooseArgs[$chosenEntry]);
        }

        return $chooseArgs[$chosenEntry];
    }

  /**
   * CHOOSECOLS.
   *
   * Returns the specified columns from an array.
   *
   * @param mixed $cells The cells being searched
   * @param int $cols List of numeric column indexes to extract
   *
   * @see https://support.microsoft.com/en-us/office/choosecols-function-bf117976-2722-4466-9b9a-1c01ed9aebff
   * @see https://support.google.com/docs/answer/13197914?hl=en
   *
   * @return array|string The resulting array, or a string containing an error
   */
  public static function choosecols(mixed $cells, int ...$cols): array|string
  {
    $columns = RowColumnInformation::COLUMNS($cells);
    if (is_string($columns)) {
      return $columns;
    }
    $result = [];
    foreach ($cols as $col) {
      if (!$col || abs($col) > $columns) {
        return ExcelError::VALUE();
      }
      $result[] = array_column($cells, $col > 0 ? $col-1 : $columns-$col);
    }
    return Matrix::transpose($result);
  }
}
