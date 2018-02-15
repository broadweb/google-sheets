<?php

namespace WhiteGrey\GoogleSheets;

class GoogleSheets
{
	/**
	 *
	 */
	private $_service;


	/**
	 * Sets config including Google credentials
	 *
	 * @param array $config
	 */
	public function __construct($pathToJsonCreds)
	{
		putenv('GOOGLE_APPLICATION_CREDENTIALS='.$pathToJsonCreds);
		$client = new \Google_Client();
		$client->useApplicationDefaultCredentials();
		$client->addScope(\Google_Service_Sheets::SPREADSHEETS);

		$this->_service = new \Google_Service_Sheets($client);
	}

	/**
	 *
	 */
	public function getService()
	{
		return $this->_service;
	}


	/**
	 * Return rows from a spreadsheet for given range
	 *
	 * @param string $spreadsheetId Google sheet ID
	 * @param string $range Range
	 *
	 * @return array Desired rows
	 */
	public function getRows($spreadsheetId, $range)
	{
		$response = $this->_service->spreadsheets_values->get($spreadsheetId, $range);

		$rows = $response->values;

		if (is_array($rows) === false || count($rows) === 0) {
			return array();
		}

		return $rows;
	}

	/**
	 *
	 */
	public function clearRange($spreadsheetId, $range)
	{
		$requestBody = new \Google_Service_Sheets_ClearValuesRequest();

		$this->_service->spreadsheets_values->clear($spreadsheetId, $range, $requestBody);
	}


	/**
	 * Write given rows to Google sheet
	 *
	 * @param string $spreadsheetId Google sheet ID
	 * @param string $cell First cell from which rows are to be written
	 * @param array $rows Rows to be written
	 *
	 * @boolean True on success, false otherwise
	 */
	public function writeRows($spreadsheetId, $cell, array $rows)
	{
		$body = new \Google_Service_Sheets_ValueRange(array(
		  'values' => $rows
		));

		$valueInputOption = 'RAW';
		$params = array(
		  'valueInputOption' => $valueInputOption
		);
		$result = $this->_service->spreadsheets_values->update($spreadsheetId, $cell, $body, $params);

		if ($result->updatedRows === count($rows)) {
			return true;
		}
		return false;
	}


	/**
	 * Write given value into given cell of Google sheet
	 *
	 * @param string $spreadsheetId Google sheet ID
	 * @param string $cell First cell where row is to be written
	 * @param array $row Row to be written
	 *
	 * @boolean True on success, false otherwise
	 */
	public function writeRow($spreadsheetId, $cell, array $row)
	{
		$values = array(
				  	$row
				  );
		$body = new \Google_Service_Sheets_ValueRange(array(
		  'values' => $values
		));

		$valueInputOption = 'RAW';
		$params = array(
		  'valueInputOption' => $valueInputOption
		);
		$result = $this->_service->spreadsheets_values->update($spreadsheetId, $cell, $body, $params);

		if ($result->updatedCells === count($row)) {
			return true;
		}
		return false;
	}


	/**
	 * Write given value into given cell of Google sheet
	 *
	 * @param string $spreadsheetId Google sheet ID
	 * @param string $cell Range that resolves to one cell
	 * @param string $value Value to be written
	 *
	 * @boolean True on success, false otherwise
	 */
	public function write($spreadsheetId, $cell, $value)
	{
		$values = array(
		    array(
		        $value,
		    ),
		);
		$body = new \Google_Service_Sheets_ValueRange(array(
		  'values' => $values
		));

		$valueInputOption = 'RAW';
		$params = array(
		  'valueInputOption' => $valueInputOption
		);
		$result = $this->_service->spreadsheets_values->update($spreadsheetId, $cell, $body, $params);

		if ($result->updatedCells == 1) {
			return true;
		}
		return false;

	}

	/**
	 * Gets column index for given column name in row
	 *
	 * @param array $row Row
	 * @param string $columnName Column name
	 *
	 * @return int|null Column index or null if column not found
	 */
    public function getColumnFor($row, $columnName)
    {
    	$index = $this->getColumnIndexFor($row, $columnName);
    	if ($index === null) {
    		return null;
    	}
    	return $this->_getNameFromNumber($index);
	}

	/**
	 * Gets column index for given column name in row
	 *
	 * @param array $row Row
	 * @param string $columnName Column name
	 *
	 * @return int|null Column index or null if column not found
	 */
    public function getColumnIndexFor($row, $columnName)
    {
    	$index = 0;
    	foreach ($row as $column)
    	{
    		if ($column === $columnName) {
    			return $index;
    		}
    		$index++;
    	}

    	return null;
	}


	/**
	 * Get column name from column index
	 *
	 * @param int $num Column index
	 *
	 * @return string $letter Column letter
     */
	private function _getNameFromNumber($num)
	{
	    $numeric = $num % 26;
	    $letter = chr(65 + $numeric);
	    $num2 = intval($num / 26);
	    if ($num2 > 0) {
	        return getNameFromNumber($num2 - 1) . $letter;
	    } else {
	        return $letter;
	    }
	}
}
