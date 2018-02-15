<?php

/**
 *
 */

use WhiteGrey\GoogleSheets\GoogleSheets;

class GoogleSheetsTest extends PHPUnit_Framework_TestCase{

  private $_sheet;

  /**
  *
  */
  public function setUp()
  {
    $creds = __DIR__.'/../google-creds.json';
    $this->_sheet = new GoogleSheets($creds);
  }


  /**
  *
  */
  public function testGoogleSheet()
  {
    $this->assertInstanceOf(Google_Service_Sheets::class, $this->_sheet->getService());
  }

  /**
  *
  */
  public function testGetRows()
  {
    $spreadsheetId = 'xxxxxxxxxxxxxxxxxx';
    $range = 'Facebook Posts!A:Z';
    $rows = $this->_sheet->getRows($spreadsheetId, $range);
    $this->assertCount(43, $rows);
  }

  /**
  *
  */
  public function testWrite()
  {
    $spreadsheetId = 'xxxxxxxxxxxxxxxxxx';
    $cell = 'Facebook Video Stats (Script)!V2';
    $result = $this->_sheet->write($spreadsheetId, $cell, 'PHPUNIT');
    $this->assertTrue($result);
  }

}