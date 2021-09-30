<?php

/**
 * A class for reading Microsoft Excel (97/2003) Spreadsheets.
*/

class Spreadsheet_Excel_Reader {

	const NUM_BIG_BLOCK_DEPOT_BLOCKS_POS = 0x2c;
	const SMALL_BLOCK_DEPOT_BLOCK_POS = 0x3c;
	const ROOT_START_BLOCK_POS = 0x30;
	const BIG_BLOCK_SIZE = 0x200;
	const SMALL_BLOCK_SIZE = 0x40;
	const EXTENSION_BLOCK_POS = 0x44;
	const NUM_EXTENSION_BLOCK_POS = 0x48;
	const PROPERTY_STORAGE_BLOCK_SIZE = 0x80;
	const BIG_BLOCK_DEPOT_BLOCKS_POS = 0x4c;
	const SMALL_BLOCK_THRESHOLD = 0x1000;
	// property storage offsets
	const SIZE_OF_NAME_POS = 0x40;
	const TYPE_POS = 0x42;
	const START_BLOCK_POS = 0x74;
	const SIZE_POS = 0x78;
	
	const BIFF8 = 0x600;
	const BIFF7 = 0x500;
	const WORKBOOKGLOBALS = 0x5;
	const WORKSHEET = 0x10;
	const TYPE_BOF = 0x809;
	const TYPE_EOF = 0x0a;
	const TYPE_BOUNDSHEET = 0x85;
	const TYPE_DIMENSION =	0x200;
	const TYPE_ROW = 0x208;
	const TYPE_DBCELL = 0xd7;
	const TYPE_FILEPASS = 0x2f;
	const TYPE_NOTE = 0x1c;
	const TYPE_TXO = 0x1b6;
	const TYPE_RK = 0x7e;
	const TYPE_RK2 = 0x27e;
	const TYPE_MULRK = 0xbd;
	const TYPE_MULBLANK = 0xbe;
	const TYPE_INDEX = 0x20b;
	const TYPE_SST = 0xfc;
	const TYPE_EXTSST = 0xff;
	const TYPE_CONTINUE = 0x3c;
	const TYPE_LABEL = 0x204;
	const TYPE_LABELSST = 0xfd;
	const TYPE_NUMBER = 0x203;
	const TYPE_NAME = 0x18;
	const TYPE_ARRAY = 0x221;
	const TYPE_STRING = 0x207;
	const TYPE_FORMULA = 0x406;
	const TYPE_FORMULA2 = 0x6;
	const TYPE_FORMAT = 0x41e;
	const TYPE_XF = 0xe0;
	const TYPE_BOOLERR = 0x205;
	const TYPE_FONT = 0x0031;
	const TYPE_PALETTE = 0x0092;
	const TYPE_UNKNOWN = 0xffff;
	const TYPE_NINETEENFOUR = 0x22;
	const TYPE_MERGEDCELLS = 0xE5;
	const UTCOFFSETDAYS = 25569;
	const UTCOFFSETDAYS1904 = 24107;
	const MSINADAY = 86400;
	const TYPE_HYPER = 0x01b8;
	const TYPE_COLINFO = 0x7d;
	const TYPE_DEFCOLWIDTH = 0x55;
	const TYPE_STANDARDWIDTH = 0x99;
	const DEF_NUM_FORMAT =	"%s";

	/**
	 * List of default date formats used by Excel
	 */
	const DATE_FORMATS = array (
		0xe => "m/d/Y",
		0xf => "M-d-Y",
		0x10 => "d-M",
		0x11 => "M-Y",
		0x12 => "h:i a",
		0x13 => "h:i:s a",
		0x14 => "H:i",
		0x15 => "H:i:s",
		0x16 => "d/m/Y H:i",
		0x2d => "i:s",
		0x2e => "H:i:s",
		0x2f => "i:s.S"
	);

	/**
	 * Default number formats used by Excel
	 */
	const NUMBER_FORMATS = array(
		0x1 => "0",
		0x2 => "0.00",
		0x3 => "#,##0",
		0x4 => "#,##0.00",
		0x5 => "\$#,##0;(\$#,##0)",
		0x6 => "\$#,##0;[Red](\$#,##0)",
		0x7 => "\$#,##0.00;(\$#,##0.00)",
		0x8 => "\$#,##0.00;[Red](\$#,##0.00)",
		0x9 => "0%",
		0xa => "0.00%",
		0xb => "0.00E+00",
		0x25 => "#,##0;(#,##0)",
		0x26 => "#,##0;[Red](#,##0)",
		0x27 => "#,##0.00;(#,##0.00)",
		0x28 => "#,##0.00;[Red](#,##0.00)",
		0x29 => "#,##0;(#,##0)",  // Not exactly
		0x2a => "\$#,##0;(\$#,##0)",  // Not exactly
		0x2b => "#,##0.00;(#,##0.00)",  // Not exactly
		0x2c => "\$#,##0.00;(\$#,##0.00)",  // Not exactly
		0x30 => "##0.0E+0"
	);

	const LINE_STYLES = array(
		0x00 => "",
		0x01 => "Thin",
		0x02 => "Medium",
		0x03 => "Dashed",
		0x04 => "Dotted",
		0x05 => "Thick",
		0x06 => "Double",
		0x07 => "Hair",
		0x08 => "Medium dashed",
		0x09 => "Thin dash-dotted",
		0x0A => "Medium dash-dotted",
		0x0B => "Thin dash-dot-dotted",
		0x0C => "Medium dash-dot-dotted",
		0x0D => "Slanted medium dash-dotted"
	);	

	const LINE_STYLE_CSS = array(
		"Thin" => "1px solid", 
		"Medium" => "2px solid", 
		"Dashed" => "1px dashed", 
		"Dotted" => "1px dotted", 
		"Thick" => "3px solid", 
		"Double" => "double", 
		"Hair" => "1px solid", 
		"Medium dashed" => "2px dashed", 
		"Thin dash-dotted" => "1px dashed", 
		"Medium dash-dotted" => "2px dashed", 
		"Thin dash-dot-dotted" => "1px dashed", 
		"Medium dash-dot-dotted" => "2px dashed", 
		"Slanted medium dash-dotte" => "2px dashed" 
	);

	private $Colors = Array(
        0x00 => "#000000",
        0x01 => "#FFFFFF",
        0x02 => "#FF0000",
        0x03 => "#00FF00",
        0x04 => "#0000FF",
        0x05 => "#FFFF00",
        0x06 => "#FF00FF",
        0x07 => "#00FFFF",
        0x08 => "#000000",
        0x09 => "#FFFFFF",
        0x0A => "#FF0000",
        0x0B => "#00FF00",
        0x0C => "#0000FF",
        0x0D => "#FFFF00",
        0x0E => "#FF00FF",
        0x0F => "#00FFFF",
        0x10 => "#800000",
        0x11 => "#008000",
        0x12 => "#000080",
        0x13 => "#808000",
        0x14 => "#800080",
        0x15 => "#008080",
        0x16 => "#C0C0C0",
        0x17 => "#808080",
        0x18 => "#9999FF",
        0x19 => "#993366",
        0x1A => "#FFFFCC",
        0x1B => "#CCFFFF",
        0x1C => "#660066",
        0x1D => "#FF8080",
        0x1E => "#0066CC",
        0x1F => "#CCCCFF",
        0x20 => "#000080",
        0x21 => "#FF00FF",
        0x22 => "#FFFF00",
        0x23 => "#00FFFF",
        0x24 => "#800080",
        0x25 => "#800000",
        0x26 => "#008080",
        0x27 => "#0000FF",
        0x28 => "#00CCFF",
        0x29 => "#CCFFFF",
        0x2A => "#CCFFCC",
        0x2B => "#FFFF99",
        0x2C => "#99CCFF",
        0x2D => "#FF99CC",
        0x2E => "#CC99FF",
        0x2F => "#FFCC99",
        0x30 => "#3366FF",
        0x31 => "#33CCCC",
        0x32 => "#99CC00",
        0x33 => "#FFCC00",
        0x34 => "#FF9900",
        0x35 => "#FF6600",
        0x36 => "#666699",
        0x37 => "#969696",
        0x38 => "#003366",
        0x39 => "#339966",
        0x3A => "#003300",
        0x3B => "#333300",
        0x3C => "#993300",
        0x3D => "#993366",
        0x3E => "#333399",
        0x3F => "#333333",
        0x40 => "#000000",
        0x41 => "#FFFFFF",

        0x43 => "#000000",
        0x4D => "#000000",
        0x4E => "#FFFFFF",
        0x4F => "#000000",
        0x50 => "#FFFFFF",
        0x51 => "#000000",

        0x7FFF => "#000000"
    );

	public $colnames = array();
	public $colindexes = array();
	public $standardColWidth = 0;
	public $defaultColWidth = 0;

	public $boundsheets = array();
	public $formatRecords = array();
	public $fontRecords = array();
	public $xfRecords = array();
	public $colInfo = array();
	public $rowInfo = array();
	
	public $sst = array();
	public $sheets = array();

	private $fileData = null;
	private $data;

	private $defaultEncoding = "UTF-8";
	private $defaultFormat = self::DEF_NUM_FORMAT;
	private $columnsFormat = array();
	private $rowOffset = 1;
	private $colOffset = 1;
	private $storeExtendedInfo = true;

	function __construct(string $Filepath = '', bool $storeExtendedInfo = true, string $outputEncoding = '')
	{
		$this->setUTFEncoder('iconv');

		if ($outputEncoding != '') { 
			$this->setOutputEncoding($outputEncoding);
		}

		for ($i=1; $i<245; $i++) {
			$name = strtolower(( (($i-1)/26>=1)?chr(($i-1)/26+64):'') . chr(($i-1)%26+65));
			$this->colnames[$name] = $i;
			$this->colindexes[$i] = $name;
		}

		$this->storeExtendedInfo = $storeExtendedInfo;

		$this->read($Filepath);

		$this->data = $this->getWorkBook();

		$this->parse();
	}

	/**
	 * Set the encoding method
	 */
	function setOutputEncoding(string $encoding) : void
	{
		$this->defaultEncoding = $encoding;
	}

	/**
	 *  $encoder = 'iconv' or 'mb'
	 *  set iconv if you would like use 'iconv' for encode UTF-16LE to your encoding
	 *  set mb if you would like use 'mb_convert_encoding' for encode UTF-16LE to your encoding
	 */
	function setUTFEncoder(string $encoder = 'iconv') : void
	{
		$this->_encoderFunction = '';

		if ($encoder == 'iconv') {
			$this->_encoderFunction = function_exists('iconv') ? 'iconv' : '';
		} elseif ($encoder == 'mb') {
			$this->_encoderFunction = function_exists('mb_convert_encoding') ? 'mb_convert_encoding' : '';
		}
	}

	/**
	 * Set the row and col offset
	 */
	function setRowColOffset(int $offset) : void
	{
		$this->rowOffset = $offset;
		$this->colOffset = $offset;
	}

	/**
	 * Set the default number format
	 */
	function setDefaultFormat(string $format) : void
	{
		$this->defaultFormat = $format;
	}

	/**
	 * Force a column to use a certain format
	 */
	function setColumnFormat(int $column, string $format) : void
	{
		$this->columnsFormat[$column] = $format;
	}

	/**
	 * Read the spreadsheet file, then parse
	 */
	private function read(string $Filepath) : void
	{
		$IDENTIFIER_OLE = pack("CCCCCCCC",0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1);

		if (! is_readable($Filepath)) {
			throw new Exception(__CLASS__ . ': File not readable (' . $Filepath . ')');
		}

		$this->fileData = @file_get_contents($Filepath);

		if (! $this->fileData) {
			throw new Exception(__CLASS__ . ': No data (' . $Filepath . ')');
   		}

   		if (substr($this->fileData, 0, 8) != $IDENTIFIER_OLE) {
			throw new Exception(__CLASS__ . ': Bad identifier (' . $Filepath . ')');
   		}

		$this->numBigBlockDepotBlocks = $this->getInt4d($this->fileData, self::NUM_BIG_BLOCK_DEPOT_BLOCKS_POS);
		$this->sbdStartBlock = $this->getInt4d($this->fileData, self::SMALL_BLOCK_DEPOT_BLOCK_POS);
		$this->rootStartBlock = $this->getInt4d($this->fileData, self::ROOT_START_BLOCK_POS);
		$this->extensionBlock = $this->getInt4d($this->fileData, self::EXTENSION_BLOCK_POS);
		$this->numExtensionBlocks = $this->getInt4d($this->fileData, self::NUM_EXTENSION_BLOCK_POS);

		$bigBlockDepotBlocks = array();
		$pos = self::BIG_BLOCK_DEPOT_BLOCKS_POS;
		$bbdBlocks = $this->numBigBlockDepotBlocks;

		if ($this->numExtensionBlocks != 0) {
			$bbdBlocks = (self::BIG_BLOCK_SIZE - self::BIG_BLOCK_DEPOT_BLOCKS_POS)/4;
		}

		for ($i = 0; $i < $bbdBlocks; $i++) {
			$bigBlockDepotBlocks[$i] = $this->getInt4d($this->fileData, $pos);
			$pos += 4;
		}

		for ($j = 0; $j < $this->numExtensionBlocks; $j++) {
			$pos = ($this->extensionBlock + 1) * self::BIG_BLOCK_SIZE;
			$blocksToRead = min($this->numBigBlockDepotBlocks - $bbdBlocks, self::BIG_BLOCK_SIZE / 4 - 1);

			for ($i = $bbdBlocks; $i < $bbdBlocks + $blocksToRead; $i++) {
				$bigBlockDepotBlocks[$i] = $this->getInt4d($this->fileData, $pos);
				$pos += 4;
			}

			$bbdBlocks += $blocksToRead;
			if ($bbdBlocks < $this->numBigBlockDepotBlocks) {
				$this->extensionBlock = $this->getInt4d($this->fileData, $pos);
			}
		}

		// readBigBlockDepot
		$pos = 0;
		$index = 0;
		$this->bigBlockChain = array();

		for ($i = 0; $i < $this->numBigBlockDepotBlocks; $i++) {
			$pos = ($bigBlockDepotBlocks[$i] + 1) * self::BIG_BLOCK_SIZE;
			//echo "pos = $pos";
			for ($j = 0 ; $j < self::BIG_BLOCK_SIZE / 4; $j++) {
				$this->bigBlockChain[$index] = $this->getInt4d($this->fileData, $pos);
				$pos += 4 ;
				$index++;
			}
		}

		// readSmallBlockDepot();
		$pos = 0;
		$index = 0;
		$sbdBlock = $this->sbdStartBlock;
		$this->smallBlockChain = array();

		while ($sbdBlock != -2) {
            $pos = ($sbdBlock + 1) * self::BIG_BLOCK_SIZE;

            for ($j = 0; $j < self::BIG_BLOCK_SIZE / 4; $j++) {
                $this->smallBlockChain[$index] = $this->getInt4d($this->fileData, $pos);
                $pos += 4;
                $index++;
            }

            $sbdBlock = $this->bigBlockChain[$sbdBlock];
		}


		// readData(rootStartBlock)
		$block = $this->rootStartBlock;
		$pos = 0;
		$this->entry = $this->readData($block);
		$this->readPropertySets();
	}

	private function readData(int $bl) : string
	{
		$block = $bl;
		$pos = 0;
		$data = '';

		while ($block != -2)  {
			$pos = ($block + 1) * self::BIG_BLOCK_SIZE;
			$data = $data.substr($this->fileData, $pos, self::BIG_BLOCK_SIZE);
			$block = $this->bigBlockChain[$block];
		}

		return $data;
	}

	private function readPropertySets() : void
	{
		$offset = 0;

		while ($offset < strlen($this->entry)) {
			$d = substr($this->entry, $offset, self::PROPERTY_STORAGE_BLOCK_SIZE);
			$nameSize = ord($d[self::SIZE_OF_NAME_POS]) | (ord($d[self::SIZE_OF_NAME_POS+1]) << 8);
			$type = ord($d[self::TYPE_POS]);
			$startBlock = $this->getInt4d($d, self::START_BLOCK_POS);
			$size = $this->getInt4d($d, self::SIZE_POS);
			$name = '';

			for ($i = 0; $i < $nameSize ; $i++) {
				$name .= $d[$i];
			}

			$name = str_replace("\x00", "", $name);
			$this->props[] = array (
				'name' => $name,
				'type' => $type,
				'startBlock' => $startBlock,
				'size' => $size);

			if ((strtolower($name) == "workbook") || ( strtolower($name) == "book")) {
				$this->wrkbook = count($this->props) - 1;
			}

			if ($name == "Root Entry") {
				$this->rootentry = count($this->props) - 1;
			}

			$offset += self::PROPERTY_STORAGE_BLOCK_SIZE;
		}
	}

	function getWorkBook() : string
	{
		if ($this->props[$this->wrkbook]['size'] < self::SMALL_BLOCK_THRESHOLD) {
			$rootdata = $this->readData($this->props[$this->rootentry]['startBlock']);
			$streamData = '';
			$block = $this->props[$this->wrkbook]['startBlock'];
			$pos = 0;

			while ($block != -2) {
                $pos = $block * self::SMALL_BLOCK_SIZE;
                $streamData .= substr($rootdata, $pos, self::SMALL_BLOCK_SIZE);
                $block = $this->smallBlockChain[$block];
			}
		} else {
			$numBlocks = $this->props[$this->wrkbook]['size'] / self::BIG_BLOCK_SIZE;

			if ($this->props[$this->wrkbook]['size'] % self::BIG_BLOCK_SIZE != 0) {
				$numBlocks++;
			}

			if ($numBlocks == 0) {
				return '';
			}

			$streamData = '';
			$block = $this->props[$this->wrkbook]['startBlock'];
			$pos = 0;

			while ($block != -2) {
                $pos = ($block + 1) * self::BIG_BLOCK_SIZE;
                $streamData .= substr($this->fileData, $pos, self::BIG_BLOCK_SIZE);
                $block = $this->bigBlockChain[$block];
			}
		}

		$this->fileData = null;

		return $streamData;
	}

	/**
	 * Parse a workbook
	 */
	private function parse() : bool
	{
		$pos = 0;
		$data = $this->data;

		$code = $this->v($data,$pos);
		$length = $this->v($data,$pos+2);
		$version = $this->v($data,$pos+4);
		$substreamType = $this->v($data,$pos+6);

		$this->version = $version;

		if (($version != self::BIFF8) &&
			($version != self::BIFF7)) {
			return false;
		}

		if ($substreamType != self::WORKBOOKGLOBALS) {
			return false;
		}

		$pos += $length + 4;

		$code = $this->v($data,$pos);
		$length = $this->v($data,$pos+2);

		while ($code != self::TYPE_EOF) {
			switch ($code) {
				case self::TYPE_SST:
					$spos = $pos + 4;
					$limitpos = $spos + $length;
					$uniqueStrings = $this->getInt4d($data, $spos+4);
					$spos += 8;

					for ($i = 0; $i < $uniqueStrings; $i++) {
						// Read in the number of characters
						if ($spos == $limitpos) {
							$opcode = $this->v($data,$spos);
							$conlength = $this->v($data,$spos+2);

							if ($opcode != 0x3c) {
								return false;
							}

							$spos += 4;
							$limitpos = $spos + $conlength;
						}

						$numChars = ord($data[$spos]) | (ord($data[$spos+1]) << 8);
						$spos += 2;
						$optionFlags = ord($data[$spos]);
						$spos++;
						$asciiEncoding = (($optionFlags & 0x01) == 0) ;
						$extendedString = ( ($optionFlags & 0x04) != 0);

						// See if string contains formatting information
						$richString = ( ($optionFlags & 0x08) != 0);

						if ($richString) {
							// Read in the crun
							$formattingRuns = $this->v($data,$spos);
							$spos += 2;
						}

						if ($extendedString) {
							// Read in cchExtRst
							$extendedRunLength = $this->getInt4d($data, $spos);
							$spos += 4;
						}

						$len = ($asciiEncoding)? $numChars : $numChars*2;

						if ($spos + $len < $limitpos) {
							$retstr = substr($data, $spos, $len);
							$spos += $len;
						} else{
							// found countinue
							$retstr = substr($data, $spos, $limitpos - $spos);
							$bytesRead = $limitpos - $spos;
							$charsLeft = $numChars - (($asciiEncoding) ? $bytesRead : ($bytesRead / 2));
							$spos = $limitpos;

							while ($charsLeft > 0) {
								$opcode = $this->v($data,$spos);
								$conlength = $this->v($data,$spos+2);

								if ($opcode != 0x3c) {
									return false;
								}

								$spos += 4;
								$limitpos = $spos + $conlength;
								$option = ord($data[$spos]);
								$spos += 1;

								if ($asciiEncoding && ($option == 0)) {
									$len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
									$retstr .= substr($data, $spos, $len);
									$charsLeft -= $len;
									$asciiEncoding = true;
								} elseif (!$asciiEncoding && ($option != 0)) {
									$len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
									$retstr .= substr($data, $spos, $len);
									$charsLeft -= $len/2;
									$asciiEncoding = false;
								} elseif (!$asciiEncoding && ($option == 0)) {
									// Bummer - the string starts off as Unicode, but after the
									// continuation it is in straightforward ASCII encoding
									$len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);

									for ($j = 0; $j < $len; $j++) {
										$retstr .= $data[$spos + $j].chr(0);
									}

									$charsLeft -= $len;
									$asciiEncoding = false;
								} else {
									$newstr = '';

									for ($j = 0; $j < strlen($retstr); $j++) {
										$newstr = $retstr[$j].chr(0);
									}

									$retstr = $newstr;
									$len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
									$retstr .= substr($data, $spos, $len);
									$charsLeft -= $len/2;
									$asciiEncoding = false;
								}

								$spos += $len;
							}
						}

						$retstr = ($asciiEncoding) ? $retstr : $this->encodeUTF16($retstr);

						if ($richString) {
							$spos += 4 * $formattingRuns;
						}

						// For extended strings, skip over the extended string data
						if ($extendedString) {
							$spos += $extendedRunLength;
						}

						$this->sst[]=$retstr;
					}
					break;
				case self::TYPE_FILEPASS:
					return false;
					break;
				case self::TYPE_NAME:
					break;
				case self::TYPE_FORMAT:
					$indexCode = $this->v($data,$pos+4);
					if ($version == self::BIFF8) {
						$numchars = $this->v($data,$pos+6);
						
						if (ord($data[$pos+8]) == 0){
							$formatString = substr($data, $pos+9, $numchars);
						} else {
							$formatString = substr($data, $pos+9, $numchars*2);
						}
					} else {
						$numchars = ord($data[$pos+6]);
						$formatString = substr($data, $pos+7, $numchars*2);
					}
					
					$this->formatRecords[$indexCode] = $formatString;
					break;
				case self::TYPE_FONT:
					$height = $this->v($data,$pos+4);
					$option = $this->v($data,$pos+6);
					$color = $this->v($data,$pos+8);
					$weight = $this->v($data,$pos+10);
					$under  = ord($data[$pos+14]);
					$font = "";
					// Font name
					$numchars = ord($data[$pos+18]);

					if ((ord($data[$pos+19]) & 1) == 0) {
						$font = substr($data, $pos+20, $numchars);
					} else {
						$font = substr($data, $pos+20, $numchars*2);
						$font =  $this->encodeUTF16($font); 
					}

					$this->fontRecords[] = array(
						'height' => $height / 20,
						'italic' => !!($option & 2),
						'color' => $color,
						'under' => !($under==0),
						'bold' => ($weight==700),
						'font' => $font,
						'raw' => $this->dumpHexData($data, $pos+3, $length)
					);
					break;

				case self::TYPE_PALETTE:
					$colors = ord($data[$pos+4]) | ord($data[$pos+5]) << 8;

					for ($coli = 0; $coli < $colors; $coli++) {
						$colOff = $pos + 2 + ($coli * 4);
						$colr = ord($data[$colOff]);
						$colg = ord($data[$colOff+1]);
						$colb = ord($data[$colOff+2]);

						$this->Colors[0x07 + $coli] = '#' . $this->myhex($colr) . $this->myhex($colg) . $this->myhex($colb);
					}
					break;

				case self::TYPE_XF:
					$fontIndexCode = (ord($data[$pos+4]) | ord($data[$pos+5]) << 8) - 1;
					$fontIndexCode = max(0,$fontIndexCode);
					$indexCode = ord($data[$pos+6]) | ord($data[$pos+7]) << 8;
					$alignbit = ord($data[$pos+10]) & 3;
					$bgi = (ord($data[$pos+22]) | ord($data[$pos+23]) << 8) & 0x3FFF;
					$bgcolor = ($bgi & 0x7F);
					// $bgcolor = ($bgi & 0x3f80) >> 7;
					$align = "";

					if ($alignbit==3) { $align="right"; }
					if ($alignbit==2) { $align="center"; }

					$fillPattern = (ord($data[$pos+21]) & 0xFC) >> 2;

					if ($fillPattern == 0) {
						$bgcolor = "";
					}

					$xf = array();
					$xf['formatIndex'] = $indexCode;
					$xf['align'] = $align;
					$xf['fontIndex'] = $fontIndexCode;
					$xf['bgColor'] = $bgcolor;
					$xf['fillPattern'] = $fillPattern;

					$border = ord($data[$pos+14]) | (ord($data[$pos+15]) << 8) | (ord($data[$pos+16]) << 16) | (ord($data[$pos+17]) << 24);
					$xf['borderLeft'] = self::LINE_STYLES[($border & 0xF)];
					$xf['borderRight'] = self::LINE_STYLES[($border & 0xF0) >> 4];
					$xf['borderTop'] = self::LINE_STYLES[($border & 0xF00) >> 8];
					$xf['borderBottom'] = self::LINE_STYLES[($border & 0xF000) >> 12];
					
					$xf['borderLeftColor'] = ($border & 0x7F0000) >> 16;
					$xf['borderRightColor'] = ($border & 0x3F800000) >> 23;
					$border = (ord($data[$pos+18]) | ord($data[$pos+19]) << 8);

					$xf['borderTopColor'] = ($border & 0x7F);
					$xf['borderBottomColor'] = ($border & 0x3F80) >> 7;
											
					if (array_key_exists($indexCode, self::DATE_FORMATS)) {
						$xf['type'] = 'date';
						$xf['format'] = self::DATE_FORMATS[$indexCode];

						if ($align=='') { $xf['align'] = 'right'; }
					} elseif (array_key_exists($indexCode, self::NUMBER_FORMATS)) {
						$xf['type'] = 'number';
						$xf['format'] = self::NUMBER_FORMATS[$indexCode];

						if ($align=='') { $xf['align'] = 'right'; }
					} else {
						$isdate = FALSE;
						$formatstr = '';
						
						if ($indexCode > 0) {
							if (isset($this->formatRecords[$indexCode])) {
								$formatstr = $this->formatRecords[$indexCode];
							}

							if ($formatstr!="") {
								$tmp = preg_replace("/\;.*/","",$formatstr);
								$tmp = preg_replace("/^\[[^\]]*\]/","",$tmp);

								if (preg_match("/[^hmsday\/\-:\s\\\,AMP]/i", $tmp) == 0) { // found day and time format
									$isdate = TRUE;
									$formatstr = $tmp;
									$formatstr = str_replace(array('AM/PM','mmmm','mmm'), array('a','F','M'), $formatstr);
									// m/mm are used for both minutes and months - oh SNAP!
									// This mess tries to fix for that.
									// 'm' == minutes only if following h/hh or preceding s/ss
									$formatstr = preg_replace("/(h:?)mm?/","$1i", $formatstr);
									$formatstr = preg_replace("/mm?(:?s)/","i$1", $formatstr);
									// A single 'm' = n in PHP
									$formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
									$formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
									// else it's months
									$formatstr = str_replace('mm', 'm', $formatstr);
									// Convert single 'd' to 'j'
									$formatstr = preg_replace("/(^|[^d])d([^d]|$)/", '$1j$2', $formatstr);
									$formatstr = str_replace(array('dddd','ddd','dd','yyyy','yy','hh','h'), array('l','D','d','Y','y','H','g'), $formatstr);
									$formatstr = preg_replace("/ss?/", 's', $formatstr);
								}
							}
						}

						if ($isdate) {
							$xf['type'] = 'date';
							$xf['format'] = $formatstr;
							if ($align=='') { $xf['align'] = 'right'; }
						} else {
							// If the format string has a 0 or # in it, we'll assume it's a number
							if (preg_match("/[0#]/", $formatstr)) {
								$xf['type'] = 'number';
								if ($align=='') { $xf['align']='right'; }
							} else {
								$xf['type'] = 'other';
							}

							$xf['format'] = $formatstr;
							$xf['code'] = $indexCode;
						}
					}

					$this->xfRecords[] = $xf;
					break;
				case self::TYPE_NINETEENFOUR:
					$this->nineteenFour = (ord($data[$pos+4]) == 1);
					break;
				case self::TYPE_BOUNDSHEET:
					$rec_offset = $this->getInt4d($data, $pos+4);
					$rec_typeFlag = ord($data[$pos+8]);
					$rec_visibilityFlag = ord($data[$pos+9]);
					$rec_length = ord($data[$pos+10]);

					if ($version == self::BIFF8) {
						$chartype =  ord($data[$pos+11]);

						if ($chartype == 0) {
							$rec_name	= substr($data, $pos+12, $rec_length);
						} else {
							$rec_name	= $this->encodeUTF16(substr($data, $pos+12, $rec_length*2));
						}
					} elseif ($version == self::BIFF7) {
						$rec_name	= substr($data, $pos+11, $rec_length);
					}
					$this->boundsheets[] = array('name'=>$rec_name,'offset'=>$rec_offset);
					break;

			}

			$pos += $length + 4;
			$code = ord($data[$pos]) | ord($data[$pos+1])<<8;
			$length = ord($data[$pos+2]) | ord($data[$pos+3])<<8;
		}

		foreach ($this->boundsheets as $key=>$val){
			$this->sn = $key;
			$this->parseSheet($val['offset']);
		}

		return true;
	}

	/**
	 * Parse a worksheet
	 */
	private function parseSheet(int $spos) : int
	{
		$cont = true;
		$data = $this->data;
		// read BOF
		$code = ord($data[$spos]) | ord($data[$spos+1])<<8;
		$length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;

		$version = ord($data[$spos + 4]) | ord($data[$spos + 5])<<8;
		$substreamType = ord($data[$spos + 6]) | ord($data[$spos + 7])<<8;

		if (($version != self::BIFF8) && ($version != self::BIFF7)) {
			return -1;
		}

		if ($substreamType != self::WORKSHEET) {
			return -2;
		}

		$spos += $length + 4;

		while($cont) {
			$lowcode = ord($data[$spos]);

			if ($lowcode == self::TYPE_EOF) {
				break;
			}

			$code = $lowcode | ord($data[$spos+1])<<8;
			$length = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
			$spos += 4;
			$this->sheets[$this->sn]['maxrow'] = $this->rowOffset - 1;
			$this->sheets[$this->sn]['maxcol'] = $this->colOffset - 1;
			unset($this->rectype);

			$previousRow = 0;
			$previousCol = 0;

			switch ($code) {
				case self::TYPE_DIMENSION:
					if (!isset($this->numRows)) {
						if (($length == 10) ||  ($version == self::BIFF7)) {
							$this->sheets[$this->sn]['numRows'] = ord($data[$spos+2]) | ord($data[$spos+3]) << 8;
							$this->sheets[$this->sn]['numCols'] = ord($data[$spos+6]) | ord($data[$spos+7]) << 8;
						} else {
							$this->sheets[$this->sn]['numRows'] = ord($data[$spos+4]) | ord($data[$spos+5]) << 8;
							$this->sheets[$this->sn]['numCols'] = ord($data[$spos+10]) | ord($data[$spos+11]) << 8;
						}
					}
					break;
				case self::TYPE_MERGEDCELLS:
					$cellRanges = ord($data[$spos]) | ord($data[$spos+1])<<8;

					for ($i = 0; $i < $cellRanges; $i++) {
						$fr =  ord($data[$spos + 8*$i + 2]) | ord($data[$spos + 8*$i + 3])<<8;
						$lr =  ord($data[$spos + 8*$i + 4]) | ord($data[$spos + 8*$i + 5])<<8;
						$fc =  ord($data[$spos + 8*$i + 6]) | ord($data[$spos + 8*$i + 7])<<8;
						$lc =  ord($data[$spos + 8*$i + 8]) | ord($data[$spos + 8*$i + 9])<<8;

						if ($lr - $fr > 0) {
							$this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['rowspan'] = $lr - $fr + 1;
						}

						if ($lc - $fc > 0) {
							$this->sheets[$this->sn]['cellsInfo'][$fr+1][$fc+1]['colspan'] = $lc - $fc + 1;
						}
					}
					break;
				case self::TYPE_RK:
				case self::TYPE_RK2:
					$row = ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
					$rknum = $this->getInt4d($data, $spos + 6);
					$numValue = $this->getIEEE754($rknum);
					$info = $this->getCellDetails($spos,$numValue,$column);

					$this->addCell($row, $column, $info['string'],$info);
					break;
				case self::TYPE_LABELSST:
					$row		= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column	 = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
					$xfindex	= ord($data[$spos+4]) | ord($data[$spos+5])<<8;
					$index  = $this->getInt4d($data, $spos + 6);

					$this->addCell($row, $column, $this->sst[$index], array('xfIndex'=>$xfindex) );
					break;
				case self::TYPE_MULRK:
					$row		= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$colFirst   = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
					$colLast	= ord($data[$spos + $length - 2]) | ord($data[$spos + $length - 1])<<8;
					$columns	= $colLast - $colFirst + 1;
					$tmppos = $spos+4;

					for ($i = 0; $i < $columns; $i++) {
						$numValue = $this->getIEEE754($this->getInt4d($data, $tmppos + 2));
						$info = $this->getCellDetails($tmppos-4,$numValue,$colFirst + $i + 1);
						$tmppos += 6;

						$this->addCell($row, $colFirst + $i, $info['string'], $info);
					}
					break;
				case self::TYPE_NUMBER:
					$row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
					$tmp = unpack("ddouble", substr($data, $spos + 6, 8)); // It machine machine dependent

					if ($this->isDate($spos)) {
						$numValue = $tmp['double'];
					} else {
						$numValue = $this->createNumber($spos);
					}
					
					$info = $this->getCellDetails($spos,$numValue,$column);
					$this->addCell($row, $column, $info['string'], $info);
					break;

				case self::TYPE_FORMULA:
				case self::TYPE_FORMULA2:
					$row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;

					if ((ord($data[$spos+6])==0) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
						//String formula. Result follows in a STRING record
						// This row/col are stored to be referenced in that record
						// http://code.google.com/p/php-excel-reader/issues/detail?id=4
						$previousRow = $row;
						$previousCol = $column;
					} elseif ((ord($data[$spos+6])==1) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
						//Boolean formula. Result is in +2; 0=false,1=true
						// http://code.google.com/p/php-excel-reader/issues/detail?id=4
                        if (ord($this->data[$spos+8])==1) {
                            $this->addCell($row, $column, "TRUE");
                        } else {
                            $this->addCell($row, $column, "FALSE");
                        }
					} elseif ((ord($data[$spos+6])==2) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
						//Error formula. Error code is in +2;
					} elseif ((ord($data[$spos+6])==3) && (ord($data[$spos+12])==255) && (ord($data[$spos+13])==255)) {
						//Formula result is a null string.
						$this->addCell($row, $column, '');
					} else {
						// result is a number, so first 14 bytes are just like a _NUMBER record
						$tmp = unpack("ddouble", substr($data, $spos + 6, 8)); // It machine machine dependent

						if ($this->isDate($spos)) {
							$numValue = $tmp['double'];
						} else {
							$numValue = $this->createNumber($spos);
						}
						
						$info = $this->getCellDetails($spos,$numValue,$column);
						$this->addCell($row, $column, $info['string'], $info);
					}
					break;
				case self::TYPE_BOOLERR:
					$row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
					$string = ord($data[$spos+6]);

					$this->addCell($row, $column, $string);
					break;
                case self::TYPE_STRING:
					// http://code.google.com/p/php-excel-reader/issues/detail?id=4
					if ($version == self::BIFF8) {
						// Unicode 16 string, like an SST record
						$xpos = $spos;
						$numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
						$xpos += 2;
						$optionFlags =ord($data[$xpos]);
						$xpos++;
						$asciiEncoding = (($optionFlags &0x01) == 0) ;
						$extendedString = (($optionFlags & 0x04) != 0);
                        // See if string contains formatting information
						$richString = (($optionFlags & 0x08) != 0);

						if ($richString) {
							// Read in the crun
							$formattingRuns =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
							$xpos += 2;
						}

						if ($extendedString) {
							// Read in cchExtRst
							$extendedRunLength =$this->getInt4d($this->data, $xpos);
							$xpos += 4;
						}

						$len = ($asciiEncoding)?$numChars : $numChars*2;
						$retstr =substr($data, $xpos, $len);
						$xpos += $len;
						$retstr = ($asciiEncoding)? $retstr : $this->encodeUTF16($retstr);
					} elseif ($version == self::BIFF7) {
						// Simple byte string
						$xpos = $spos;
						$numChars =ord($data[$xpos]) | (ord($data[$xpos+1]) << 8);
						$xpos += 2;
						$retstr =substr($data, $xpos, $numChars);
					}

					$this->addCell($previousRow, $previousCol, $retstr);
					break;
				case self::TYPE_ROW:
					$row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$rowInfo = ord($data[$spos + 6]) | ((ord($data[$spos+7]) << 8) & 0x7FFF);

					if (($rowInfo & 0x8000) > 0) {
						$rowHeight = -1;
					} else {
						$rowHeight = $rowInfo & 0x7FFF;
					}

					$rowHidden = (ord($data[$spos + 12]) & 0x20) >> 5;
					$this->rowInfo[$this->sn][$row+1] = Array('height' => $rowHeight / 20, 'hidden'=>$rowHidden );
					break;
				case self::TYPE_DBCELL:
					break;
				case self::TYPE_MULBLANK:
					$row = ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;
					$cols = ($length / 2) - 3;
					
					for ($c = 0; $c < $cols; $c++) {
						$xfindex = ord($data[$spos + 4 + ($c * 2)]) | ord($data[$spos + 5 + ($c * 2)])<<8;

						$this->addCell($row, $column + $c, "", array('xfIndex'=>$xfindex));
					}
					break;
				case self::TYPE_LABEL:
					$row	= ord($data[$spos]) | ord($data[$spos+1])<<8;
					$column = ord($data[$spos+2]) | ord($data[$spos+3])<<8;

					$this->addCell($row, $column, substr($data, $spos + 8, ord($data[$spos + 6]) | ord($data[$spos + 7])<<8));
					break;
				case self::TYPE_EOF:
					$cont = false;
					break;
				case self::TYPE_HYPER:
					//  Only handle hyperlinks to a URL
					$row	= ord($this->data[$spos]) | ord($this->data[$spos+1])<<8;
					$row2   = ord($this->data[$spos+2]) | ord($this->data[$spos+3])<<8;
					$column = ord($this->data[$spos+4]) | ord($this->data[$spos+5])<<8;
					$column2 = ord($this->data[$spos+6]) | ord($this->data[$spos+7])<<8;
					$linkdata = Array();
					$flags = ord($this->data[$spos + 28]);
					$udesc = "";
					$ulink = "";
					$uloc = 32;
					$linkdata['flags'] = $flags;

					if (($flags & 1) > 0 ) {   // is a type we understand
						//  is there a description ?
						if (($flags & 0x14) == 0x14 ) {   // has a description
							$uloc += 4;
							$descLen = ord($this->data[$spos + 32]) | ord($this->data[$spos + 33]) << 8;
							$udesc = substr($this->data, $spos + $uloc, $descLen * 2);
							$uloc += 2 * $descLen;
						}

						$ulink = $this->read16bitstring($this->data, $spos + $uloc + 20);

						if ($udesc == "") {
							$udesc = $ulink;
						}
					}

					$linkdata['desc'] = $udesc;
					$linkdata['link'] = $this->encodeUTF16($ulink);

					for ($r=$row; $r<=$row2; $r++) { 
						for ($c=$column; $c<=$column2; $c++) {
							$this->sheets[$this->sn]['cellsInfo'][$r+1][$c+1]['hyperlink'] = $linkdata;
						}
					}
					break;
				case self::TYPE_DEFCOLWIDTH:
					$this->defaultColWidth  = ord($data[$spos+4]) | ord($data[$spos+5]) << 8; 
					break;
				case self::TYPE_STANDARDWIDTH:
					$this->standardColWidth  = ord($data[$spos+4]) | ord($data[$spos+5]) << 8; 
					break;
				case self::TYPE_COLINFO:
					$colfrom = ord($data[$spos+0]) | ord($data[$spos+1]) << 8;
					$colto = ord($data[$spos+2]) | ord($data[$spos+3]) << 8;
					$cw = ord($data[$spos+4]) | ord($data[$spos+5]) << 8; 
					$cxf = ord($data[$spos+6]) | ord($data[$spos+7]) << 8; 
					$co = ord($data[$spos+8]); 

					for ($coli = $colfrom; $coli <= $colto; $coli++) {
						$this->colInfo[$this->sn][$coli+1] = Array('width' => $cw, 'xf' => $cxf, 'hidden' => ($co & 0x01), 'collapsed' => ($co & 0x1000) >> 12);
					}
					break;

				default:
					break;
			}

			$spos += $length;
		
			if (!isset($this->sheets[$this->sn]['numRows'])) {
				$this->sheets[$this->sn]['numRows'] = $this->sheets[$this->sn]['maxrow'];
			}

			if (!isset($this->sheets[$this->sn]['numCols'])) {
				$this->sheets[$this->sn]['numCols'] = $this->sheets[$this->sn]['maxcol'];
			}
		}

		return 0;
	}

	/**
	 * Get the details for a particular cell
	 */
	private function getCellDetails(int $spos, float $numValue, int $column) : array
	{
		$xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
		$xfrecord = $this->xfRecords[$xfindex];
		$type = $xfrecord['type'];

		$format = $xfrecord['format'];
		$formatIndex = $xfrecord['formatIndex'];
		$fontIndex = $xfrecord['fontIndex'];
		$formatColor = "";
		$rectype = '';
		$string = '';
		$raw = '';

		if (isset($this->columnsFormat[$column + 1])){
			$format = $this->columnsFormat[$column + 1];
		}

		if ($type == 'date') {
			// See http://groups.google.com/group/php-excel-reader-discuss/browse_frm/thread/9c3f9790d12d8e10/f2045c2369ac79de
			$rectype = 'date';
			// Convert numeric value into a date
			$utcDays = floor($numValue - ($this->nineteenFour ? self::UTCOFFSETDAYS1904 : self::UTCOFFSETDAYS));
			$utcValue = ($utcDays) * self::MSINADAY;
			$dateinfo = getdate($utcValue);

			$raw = $numValue;
			$fractionalDay = $numValue - floor($numValue) + .0000001; // The .0000001 is to fix for php/excel fractional diffs

			$totalseconds = floor(self::MSINADAY * $fractionalDay);
			$secs = $totalseconds % 60;
			$totalseconds -= $secs;
			$hours = floor($totalseconds / (60 * 60));
			$mins = floor($totalseconds / 60) % 60;
			$string = date ($format, mktime($hours, $mins, $secs, $dateinfo["mon"], $dateinfo["mday"], $dateinfo["year"]));
		} else if ($type == 'number') {
			$rectype = 'number';
			$formatted = $this->formatValue($format, $numValue, $formatIndex);
			$string = $formatted['string'];
			$formatColor = $formatted['formatColor'];
			$raw = $numValue;
		} else {
			if ($format=="") {
				$format = $this->defaultFormat;
			}

			$rectype = 'unknown';
			$formatted = $this->formatValue($format, $numValue, $formatIndex);
			$string = $formatted['string'];
			$formatColor = $formatted['formatColor'];
			$raw = $numValue;
		}

		return array(
			'string' => $string,
			'raw' => $raw,
			'rectype' => $rectype,
			'format' => $format,
			'formatIndex' => $formatIndex,
			'fontIndex' => $fontIndex,
			'formatColor' => $formatColor,
			'xfIndex' => $xfindex
		);
	}

	// PUBLIC API FUNCTIONS
	// --------------------

	function val(int $row, int $col, int $sheet = 0) : string
	{
		$col = $this->getCol($col);

		if (array_key_exists($row,$this->sheets[$sheet]['cells']) && array_key_exists($col,$this->sheets[$sheet]['cells'][$row])) {
			return $this->sheets[$sheet]['cells'][$row][$col];
		}

		return "";
	}

	function value(int $row, int $col, int $sheet = 0) : string
	{
		return $this->val($row,$col,$sheet);
	}

	function info(int $row, int $col, string $type = '', int $sheet = 0) : string
	{
		$col = $this->getCol($col);

		if (array_key_exists('cellsInfo',$this->sheets[$sheet])
				&& array_key_exists($row,$this->sheets[$sheet]['cellsInfo'])
				&& array_key_exists($col,$this->sheets[$sheet]['cellsInfo'][$row])
				&& array_key_exists($type,$this->sheets[$sheet]['cellsInfo'][$row][$col])) {
			return $this->sheets[$sheet]['cellsInfo'][$row][$col][$type];
		}

		return "";
	}

	function type(int $row, int $col, int $sheet = 0) : string
	{
		return $this->info($row,$col,'type',$sheet);
	}

	function raw(int $row, int $col, int $sheet = 0) : string
	{
		return $this->info($row,$col,'raw',$sheet);
	}

	function rowspan(int $row, int $col, int $sheet = 0) : int
	{
		$val = $this->info($row,$col,'rowspan',$sheet);

		if ($val == "") {
			return 1;
		}

		return (int)$val;
	}

	function colspan(int $row, int $col, int $sheet = 0) : int
	{
		$val = $this->info($row,$col,'colspan',$sheet);

		if ($val == "") {
			return 1;
		}

		return (int)$val;
	}

	function hyperlink(int $row, int $col, int $sheet = 0) : string
	{
		$link = $this->sheets[$sheet]['cellsInfo'][$row][$col]['hyperlink'];
		if ($link) {
			return $link['link'];
		}
		return '';
	}

	function rowCount(int $sheet = 0) : int
	{
		return $this->sheets[$sheet]['numRows'];
	}

	function colCount(int $sheet = 0) : int
	{
		return $this->sheets[$sheet]['numCols'];
	}

	function colWidth(int $col, int $sheet = 0) : int
	{
		// Col width is actually the width of the number 0. So we have to estimate and come close
		return $this->colInfo[$sheet][$col]['width']/9142*200; 
	}

	function colHidden(int $col, int $sheet = 0) : bool
	{
		return !!$this->colInfo[$sheet][$col]['hidden'];
	}

	function rowHeight(int $row, int $sheet = 0) : int
	{
		return $this->rowInfo[$sheet][$row]['height'];
	}

	function rowHidden(int $row, int $sheet = 0) : bool
	{
		return !!$this->rowInfo[$sheet][$row]['hidden'];
	}
	
	// GET THE CSS FOR FORMATTING
	// ==========================
	function style(int $row, int $col, int $sheet = 0, string $properties = '') : string
	{
		$css = "";
		$font=$this->font($row,$col,$sheet);

		if ($font!="") {
			$css .= "font-family:$font;";
		}

		$align=$this->align($row,$col,$sheet);

		if ($align!="") {
			$css .= "text-align:$align;";
		}

		$height=$this->height($row,$col,$sheet);

		if ($height!="") {
			$css .= "font-size:$height"."px;";
		}

		$bgcolor=$this->bgColor($row,$col,$sheet);

		if ($bgcolor!="") {
			$bgcolor = $this->Colors[$bgcolor];
			$css .= "background-color:$bgcolor;";
		}

		$color=$this->color($row,$col,$sheet);

		if ($color!="") {
			$css .= "color:$color;";
		}

		$bold=$this->bold($row,$col,$sheet);

		if ($bold) {
			$css .= "font-weight:bold;";
		}

		$italic=$this->italic($row,$col,$sheet);

		if ($italic) {
			$css .= "font-style:italic;";
		}

		$underline = $this->underline($row,$col,$sheet);

		if ($underline) {
			$css .= "text-decoration:underline;";
		}
		// Borders
		$bLeft = $this->borderLeft($row,$col,$sheet);
		$bRight = $this->borderRight($row,$col,$sheet);
		$bTop = $this->borderTop($row,$col,$sheet);
		$bBottom = $this->borderBottom($row,$col,$sheet);
		$bLeftCol = $this->borderLeftColor($row,$col,$sheet);
		$bRightCol = $this->borderRightColor($row,$col,$sheet);
		$bTopCol = $this->borderTopColor($row,$col,$sheet);
		$bBottomCol = $this->borderBottomColor($row,$col,$sheet);
		// Try to output the minimal required style
		if ($bLeft!="" && $bLeft==$bRight && $bRight==$bTop && $bTop==$bBottom) {
			$css .= "border:" . self::LINE_STYLE_CSS[$bLeft] .";";
		} else {
			if ($bLeft!="") { $css .= "border-left:" . self::LINE_STYLE_CSS[$bLeft] .";"; }
			if ($bRight!="") { $css .= "border-right:" . self::LINE_STYLE_CSS[$bRight] .";"; }
			if ($bTop!="") { $css .= "border-top:" . self::LINE_STYLE_CSS[$bTop] .";"; }
			if ($bBottom!="") { $css .= "border-bottom:" . self::LINE_STYLE_CSS[$bBottom] .";"; }
		}
		// Only output border colors if there is an actual border specified
		if ($bLeft!="" && $bLeftCol!="") { $css .= "border-left-color:" . $bLeftCol .";"; }
		if ($bRight!="" && $bRightCol!="") { $css .= "border-right-color:" . $bRightCol .";"; }
		if ($bTop!="" && $bTopCol!="") { $css .= "border-top-color:" . $bTopCol . ";"; }
		if ($bBottom!="" && $bBottomCol!="") { $css .= "border-bottom-color:" . $bBottomCol .";"; }
		
		return $css;
	}
	
	// FORMAT PROPERTIES
	// =================
	function format(int $row, int $col, int $sheet = 0) : string
	{
		return $this->info($row,$col,'format',$sheet);
	}

	function formatIndex(int $row, int $col, int $sheet = 0) : string
	{
		return $this->info($row,$col,'formatIndex',$sheet);
	}

	function formatColor(int $row, int $col, int $sheet = 0) : string
	{
		return $this->info($row,$col,'formatColor',$sheet);
	}
	
	// CELL (XF) PROPERTIES
	// ====================
	function xfRecord(int $row, int $col, int $sheet = 0) : ?string
	{
		$xfIndex = $this->info($row,$col,'xfIndex',$sheet);

		if ($xfIndex != "") {
			return $this->xfRecords[$xfIndex];
		}

		return null;
	}

	function xfProperty(int $row, int $col, int $sheet, string $prop) : string
	{
		$xfRecord = $this->xfRecord($row,$col,$sheet);

		if ($xfRecord != null) {
			return $xfRecord[$prop];
		}

		return "";
	}

	function align(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'align');
	}

	function bgColor(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'bgColor');
	}

	function borderLeft(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'borderLeft');
	}

	function borderRight(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'borderRight');
	}

	function borderTop(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'borderTop');
	}

	function borderBottom(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'borderBottom');
	}

	function borderLeftColor(int $row, int $col, int $sheet = 0) : string
	{
		return $this->Colors[$this->xfProperty($row,$col,$sheet,'borderLeftColor')];
	}

	function borderRightColor(int $row, int $col, int $sheet = 0) : string
	{
		return $this->Colors[$this->xfProperty($row,$col,$sheet,'borderRightColor')];
	}

	function borderTopColor(int $row, int $col, int $sheet = 0) : string
	{
		return $this->Colors[$this->xfProperty($row,$col,$sheet,'borderTopColor')];
	}

	function borderBottomColor(int $row, int $col, int $sheet = 0) : string
	{
		return $this->Colors[$this->xfProperty($row,$col,$sheet,'borderBottomColor')];
	}

	// FONT PROPERTIES
	// ===============
	function fontRecord(int $row, int $col, int $sheet = 0) : ?string
	{
	    $xfRecord = $this->xfRecord($row,$col,$sheet);

		if ($xfRecord!=null) {
			$font = $xfRecord['fontIndex'];

			if ($font != null) {
				return $this->fontRecords[$font];
			}
		}

		return null;
	}

	function fontProperty(int $row, int $col, int $sheet = 0, string $prop) : ?string
	{
		$font = $this->fontRecord($row,$col,$sheet);

		if ($font != null) {
			return $font[$prop];
		}

		return null;
	}

	function fontIndex(int $row, int $col, int $sheet = 0) : string
	{
		return $this->xfProperty($row,$col,$sheet,'fontIndex');
	}

	function color(int $row, int $col, int $sheet = 0) : string
	{
		$formatColor = $this->formatColor($row,$col,$sheet);

		if ($formatColor != "") {
			return $formatColor;
		}

		$ci = $this->fontProperty($row,$col,$sheet,'color');
        
		return $this->rawColor($ci);
    }

	function rawColor(string $ci) : string
	{
		if (($ci <> 0x7FFF) && ($ci <> '')) {
			return $this->Colors[$ci];
		}

		return "";
	}

	function bold(int $row, int $col, int $sheet = 0) : ?string
	{
		return $this->fontProperty($row,$col,$sheet,'bold');
	}

	function italic(int $row, int $col, int $sheet = 0) : ?string
	{
		return $this->fontProperty($row,$col,$sheet,'italic');
	}

	function underline(int $row, int $col, int $sheet = 0) : ?string
	{
		return $this->fontProperty($row,$col,$sheet,'under');
	}

	function height(int $row, int $col, int $sheet = 0) : ?string
	{
		return $this->fontProperty($row,$col,$sheet,'height');
	}

	function font(int $row, int $col, int $sheet = 0) : ?string
	{
		return $this->fontProperty($row,$col,$sheet,'font');
	}
	
	// DUMP AN HTML TABLE OF THE ENTIRE XLS DATA
	// =========================================
	function dump(bool $row_numbers = false, bool $col_letters = false, int $sheet = 0, string $table_class = 'excel') : string
	{
		$out = "<table class=\"$table_class\" cellspacing=0>";

		if ($col_letters) {
			$out .= "<thead>\n\t<tr>";

			if ($row_numbers) {
				$out .= "\n\t\t<th>&nbsp</th>";
			}

			for($i=1;$i<=$this->colCount($sheet);$i++) {
				$style = "width:" . ($this->colWidth($i,$sheet)*1) . "px;";

				if ($this->colHidden($i,$sheet)) {
					$style .= "display:none;";
				}

				$out .= "\n\t\t<th style=\"$style\">" . strtoupper($this->colindexes[$i]) . "</th>";
			}

			$out .= "</tr></thead>\n";
		}
		
		$out .= "<tbody>\n";

		for($row=1;$row<=$this->rowCount($sheet);$row++) {
			$rowheight = $this->rowHeight($row,$sheet);
			$style = "height:" . ($rowheight*(4/3)) . "px;";

			if ($this->rowHidden($row,$sheet)) {
				$style .= "display:none;";
			}

			$out .= "\n\t<tr style=\"$style\">";

			if ($row_numbers) {
				$out .= "\n\t\t<th>$row</th>";
			}

			for($col=1;$col<=$this->colCount($sheet);$col++) {
				// Account for Rowspans/Colspans
				$rowspan = $this->rowspan($row,$col,$sheet);
				$colspan = $this->colspan($row,$col,$sheet);

				for($i=0;$i<$rowspan;$i++) {
					for($j=0;$j<$colspan;$j++) {
						if ($i>0 || $j>0) {
							$this->sheets[$sheet]['cellsInfo'][$row+$i][$col+$j]['dontprint']=1;
						}
					}
				}

				if(!$this->sheets[$sheet]['cellsInfo'][$row][$col]['dontprint']) {
					$style = $this->style($row,$col,$sheet);

					if ($this->colHidden($col,$sheet)) {
						$style .= "display:none;";
					}

					$out .= "\n\t\t<td style=\"$style\"" . ($colspan > 1?" colspan=$colspan":"") . ($rowspan > 1?" rowspan=$rowspan":"") . ">";
					$val = $this->val($row,$col,$sheet);

					if ($val=='') {
						$val="&nbsp;";
					} else { 
						$val = htmlentities($val); 
						$link = $this->hyperlink($row,$col,$sheet);

						if ($link!='') {
							$val = "<a href=\"$link\">$val</a>";
						}
					}

					$out .= "<nobr>".nl2br($val)."</nobr>";
					$out .= "</td>";
				}
			}

			$out .= "</tr>\n";
		}

		$out .= "</tbody></table>";

		return $out;
	}
	
	// --------------
	// END PUBLIC API

	private function myHex(int $d) : string
	{
		if ($d < 16) {
			return "0" . dechex($d);
		}

		return dechex($d);
	}
	
	private function dumpHexData(string $data, int $pos, int $length) : string
	{
		$info = "";

		for ($i = 0; $i <= $length; $i++) {
			$info .= ($i==0?"":" ") . $this->myHex(ord($data[$pos + $i])) . (ord($data[$pos + $i])>31? "[" . $data[$pos + $i] . "]":'');
		}

		return $info;
	}

	private function getCol(int $col) : int
	{
		if (is_string($col)) {
			$col = strtolower($col);

			if (array_key_exists($col,$this->colnames)) {
				$col = $this->colnames[$col];
			}
		}

		return $col;
	}
	
	private function read16bitstring(string $data, int $start) : string
	{
		$len = 0;

		while (ord($data[$start + $len]) + ord($data[$start + $len + 1]) > 0) {
			$len++;
		}

		return substr($data, $start, $len);
	}
	
	private function formatValue(string $format, float $num, ?int $f) : array
	{
		// 49==TEXT format
		// http://code.google.com/p/php-excel-reader/issues/detail?id=7
		if ( (!$f && $format=="%s") || ($f==49) || ($format=="GENERAL") ) { 
			return array('string'=>$num, 'formatColor'=>null); 
		}

		// Custom pattern can be POSITIVE;NEGATIVE;ZERO
		// The "text" option as 4th parameter is not handled
		$parts = explode(";",$format);
		$pattern = $parts[0];
		// Negative pattern
		if (count($parts)>2 && $num==0) {
			$pattern = $parts[2];
		}
		// Zero pattern
		if (count($parts)>1 && $num<0) {
			$pattern = $parts[1];
			$num = abs($num);
		}

		$color = "";
		$matches = array();
		$color_regex = "/^\[(BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW)\]/i";

		if (preg_match($color_regex,$pattern,$matches)) {
			$color = strtolower($matches[1]);
			$pattern = preg_replace($color_regex,"",$pattern);
		}
		
		// In Excel formats, "_" is used to add spacing, which we can't do in HTML
		$pattern = preg_replace("/_./","",$pattern);
		
		// Some non-number characters are escaped with \, which we don't need
		$pattern = preg_replace("/\\\/","",$pattern);
		
		// Some non-number strings are quoted, so we'll get rid of the quotes
		$pattern = preg_replace("/\"/","",$pattern);

		// TEMPORARY - Convert # to 0
		$pattern = preg_replace("/\#/","0",$pattern);

		// Find out if we need comma formatting
		$has_commas = preg_match("/,/",$pattern);

		if ($has_commas) {
			$pattern = preg_replace("/,/","",$pattern);
		}

		// Handle Percentages
		if (preg_match("/\d(\%)([^\%]|$)/",$pattern,$matches)) {
			$num = $num * 100;
			$pattern = preg_replace("/(\d)(\%)([^\%]|$)/","$1%$3",$pattern);
		}

		// Handle the number itself
		$number_regex = "/(\d+)(\.?)(\d*)/";

		if (preg_match($number_regex,$pattern,$matches)) {
			$left = $matches[1];
			$dec = $matches[2];
			$right = $matches[3];

			if ($has_commas) {
				$formatted = number_format($num,strlen($right));
			} else {
				$sprintf_pattern = "%1.".strlen($right)."f";
				$formatted = sprintf($sprintf_pattern, $num);
			}

			$pattern = preg_replace($number_regex, $formatted, $pattern);
		}

		return array(
			'string'=>$pattern,
			'formatColor'=>$color
		);
	}

	private function isDate(int $spos) : bool
	{
		$xfindex = ord($this->data[$spos+4]) | ord($this->data[$spos+5]) << 8;
		return ($this->xfRecords[$xfindex]['type'] == 'date');
	}

	private function createNumber(int $spos) : float
	{
		$rknumhigh = $this->getInt4d($this->data, $spos + 10);
		$rknumlow = $this->getInt4d($this->data, $spos + 6);
		$sign = ($rknumhigh & 0x80000000) >> 31;
		$exp =  ($rknumhigh & 0x7ff00000) >> 20;
		$mantissa = (0x100000 | ($rknumhigh & 0x000fffff));
		$mantissalow1 = ($rknumlow & 0x80000000) >> 31;
		$mantissalow2 = ($rknumlow & 0x7fffffff);
		$value = $mantissa / pow( 2 , (20- ($exp - 1023)));

		if ($mantissalow1 != 0) {
			$value += 1 / pow (2 , (21 - ($exp - 1023)));
		}

		$value += $mantissalow2 / pow (2 , (52 - ($exp - 1023)));

		if ($sign) {
			$value = -1 * $value;
		}

		return  $value;
	}

	private function addCell(int $row, int $col, string $string, ?array $info = null) : void
	{
		$this->sheets[$this->sn]['maxrow'] = max($this->sheets[$this->sn]['maxrow'], $row + $this->rowOffset);
		$this->sheets[$this->sn]['maxcol'] = max($this->sheets[$this->sn]['maxcol'], $col + $this->colOffset);
		$this->sheets[$this->sn]['cells'][$row + $this->rowOffset][$col + $this->colOffset] = $string;

		if ($this->storeExtendedInfo && $info) {
			foreach ($info as $key => $val) {
				$this->sheets[$this->sn]['cellsInfo'][$row + $this->rowOffset][$col + $this->colOffset][$key] = $val;
			}
		}
	}

	private function getIEEE754(int $rknum) : float
	{
		if (($rknum & 0x02) != 0) {
			$value = $rknum >> 2;
		} else {
			//mmp
			// I got my info on IEEE754 encoding from
			// http://research.microsoft.com/~hollasch/cgindex/coding/ieeefloat.html
			// The RK format calls for using only the most significant 30 bits of the
			// 64 bit floating point value. The other 34 bits are assumed to be 0
			// So, we use the upper 30 bits of $rknum as follows...
			$sign = ($rknum & 0x80000000) >> 31;
			$exp = ($rknum & 0x7ff00000) >> 20;
			$mantissa = (0x100000 | ($rknum & 0x000ffffc));
			$value = $mantissa / pow( 2 , (20- ($exp - 1023)));

			if ($sign) {
				$value = -1 * $value;
			}
			//end of changes by mmp
		}

		if (($rknum & 0x01) != 0) {
			$value /= 100;
		}

		return $value;
	}

	private function encodeUTF16(string $string) : string
	{
		$result = $string;

		if ($this->defaultEncoding) {
			switch ($this->_encoderFunction) {
				case 'iconv' :
					$result = iconv('UTF-16LE', $this->defaultEncoding, $string);
					break;
				case 'mb_convert_encoding' :
					$result = mb_convert_encoding($string, $this->defaultEncoding, 'UTF-16LE' );
					break;
			}
		}

		return $result;
	}

	private function getInt4d(string $data, int $pos) : int
	{
		$value = ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);

		if ($value >= 4294967294) {
			$value=-2;
		}

		return $value;
	}

	private function v(string $data, int $pos)
	{
		return ord($data[$pos]) | ord($data[$pos+1])<<8;
	}

}

?>
