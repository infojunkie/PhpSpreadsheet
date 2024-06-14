<?php

namespace PhpOffice\PhpSpreadsheet\Calculation\Google;

use PhpOffice\PhpSpreadsheet\Calculation\Information\ExcelError;
use PhpOffice\PhpSpreadsheet\Calculation\Functions;
use PhpOffice\PhpSpreadsheet\Cell\Cell;

class Google
{
  /**
   * __xludf.DUMMYFUNCTION.
   * Function inserted by Google Sheets when exporting to XLSX.
   * Contains the original formula as text string.
   */
  public static function dummyfunction(string $formula, ?Cell $cell = null): mixed {
    global $calculation;
    return $calculation->calculateFormula("={$formula}", $cell->getCoordinate(), $cell);
  }

  /**
   * QUERY.
   * Runs a Google Visualization API Query Language query across data.
   *
   * @see https://support.google.com/docs/answer/3093343
   *
   * Load the dataset into a SQLite table and query it.
   */
  public static function query(array $data, $query, $headers = -1) {
    global $calculation;
    if (empty($data) || empty(reset($data))) return [];

    // Create the SQLite table.
    // The keys of each data entry are the column names.
    // TODO Detect or read explicit header rows.
    try {
      $db = new \SQLite3(':memory:');
      $columns = [];
      foreach (reset($data) as $column => $value) {
        $value = Functions::flattenSingleValue($value);
        switch (gettype($value)) {
          case 'boolean': $type = 'TINYINT'; break;
          case 'integer': $type = 'INT'; break;
          case 'double':  $type = 'REAL'; break;
          case 'string':  $type = 'TEXT'; break;
          default:
            $calculation->getDebugLog()->writeDebugLog("Evaluating SQLite3 Query: Unhandled data type %s at %s: %s", gettype($value), $column, $calculation->showValue($value));
            return ExcelError::VALUE();
        }
        $columns[$column] = $type;
      }
      $create = ('CREATE TABLE sheet(_row INTEGER PRIMARY KEY, ' . join(', ', array_map(function($c, $t) {
        return "$c $t";
      }, array_keys($columns), array_values($columns))) . ');');
      $calculation->getDebugLog()->writeDebugLog("Evaluating SQLite3 Query: %s", $create);
      $db->query($create);

      // Populate the table.
      $insert = ('INSERT INTO sheet VALUES(' . join("),\n(", array_map(function($row, $values) use ($columns) {
        return $row . ', ' . join(', ', array_map(function($column, $value) use ($columns) {
          if (is_array($value)) {
            $value = Functions::flattenSingleValue($value);
          }
          if (is_null($value)) {
            return 'NULL';
          }
          switch ($columns[$column]) {
            case 'TINYINT':
              return intval(boolval($value));
            case 'INT':
              return intval($value);
            case 'REAL':
              return doubleval($value);
            case 'TEXT':
              return "'" . \SQLite3::escapeString($value) . "'";
          }
        }, array_keys($values), array_values($values)));
      }, array_keys($data), array_values($data))) . ')');
      $calculation->getDebugLog()->writeDebugLog("Evaluating SQLite3 Query: %s", $insert);
      $db->query($insert);

      // Query the table.
      // The incoming query has no FROM clause: Add it here.
      $query = Functions::flattenSingleValue($query);
      $query = str_ireplace('WHERE', 'FROM sheet WHERE', $query);
      $result = $db->query($query);
      if ($result) {
        $rows = [];
        if ($result->numColumns() && $result->columnType(0) != SQLITE3_NULL) {
          while (FALSE !== ($row = $result->fetchArray(SQLITE3_NUM))) {
            $rows[] = $row;
          }
        }
        $calculation->getDebugLog()->writeDebugLog("Evaluating SQLite3 Query: %s", $query);
        $result = $rows;
      }
      else {
        $calculation->getDebugLog()->writeDebugLog("Evaluating SQLite3 Query: %s returned error #%d %s", $query, $db->lastErrorCode(), $db->lastErrorMsg());
        $result = ExcelError::VALUE();
      }
    }
    catch (\Exception|\TypeError $e) {
      $calculation->getDebugLog()->writeDebugLog("Evaluating SQLite3 Query: %s threw exception %s", $query, $e->getMessage());
      $result = ExcelError::VALUE();
    }
    if ($db) $db->close();
    return $result;
  }
}
