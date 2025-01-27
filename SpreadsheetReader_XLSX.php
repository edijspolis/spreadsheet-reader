<?php
/**
 * Class for parsing XLSX files specifically
 *
 * @author Martins Pilsetnieks
 */
	class SpreadsheetReader_XLSX implements Iterator, Countable
	{
		const CELL_TYPE_BOOL = 'b';
		const CELL_TYPE_NUMBER = 'n';
		const CELL_TYPE_ERROR = 'e';
		const CELL_TYPE_SHARED_STR = 's';
		const CELL_TYPE_STR = 'str';
		const CELL_TYPE_INLINE_STR = 'inlineStr';

		/**
		 * Memory used by shared strings that can be reasonably cached, i.e., that aren't read from file but stored in memory.
		 *	If the total size in bytes of shared strings is higher than this, only the allowed amount of memory will be used.
		 *	If this value is null, shared strings are cached regardless of amount.
		 *	With large shared string caches there are huge performance gains, however a lot of memory could be used which
		 *	can be a problem, especially on shared hosting.
		 */
		const SHARED_STRING_CACHE_LIMIT = 500000;

		/**
		 * Max row count of shared strings chunk file.
		 * The lower the number the more chunk files will be created.
		 * More chunk files means more XMLReader objects created to read them and more memmory used.
		 * If there is less rows in file we will need to read less lines to find the correct one and that means better performance.
		 */
		const SHARED_STRING_CHUNK_SIZE = 500;

		private $Options = array(
			'TempDir' => '',
			'ReturnDateTimeObjects' => false
		);

		private static $RuntimeInfo = array(
			'GMPSupported' => false
		);

		private $Valid = false;

		/**
		 * @var SpreadsheetReader_* Handle for the reader object
		 */
		private $Handle = false;

		// Worksheet file
		/**
		 * @var string Path to the worksheet XML file
		 */
		private $WorksheetPath = false;
		/**
		 * @var XMLReader XML reader object for the worksheet XML file
		 */
		private $Worksheet = false;

		/**
		 * @var array Collection of XML reader objects for the shared strings chunk files
		 */
        private $SharedStringReaders = array();

		/**
		 * @var array Shared strings count
		 */
		private $SharedStringCount = 0;

		/**
		 * @var array Shared strings cache, if the number of shared strings is low enough
		 */
		private $SharedStringCache = array();

		// Workbook data
		/**
		 * @var SimpleXMLElement XML object for the workbook XML file
		 */
		private $WorkbookXML = false;

		// Style data
		/**
		 * @var SimpleXMLElement XML object for the styles XML file
		 */
		private $StylesXML = false;
		/**
		 * @var array Container for cell value style data
		 */
		private $Styles = array();

		private $TempDir = '';
		private $TempFiles = array();

		private $CurrentRow = false;

		// Runtime parsing data
		/**
		 * @var int Current row in the file
		 */
		private $Index = 0;

		/**
		 * @var array Data about separate sheets in the file
		 */
		private $Sheets = false;

		private $RowOpen = false;

		private static $BuiltinFormats = array(
			0 => '',
			1 => '0',
			2 => '0.00',
			3 => '#,##0',
			4 => '#,##0.00',

			9 => '0%',
			10 => '0.00%',
			11 => '0.00E+00',
			12 => '# ?/?',
			13 => '# ??/??',
			14 => 'mm-dd-yy',
			15 => 'd-mmm-yy',
			16 => 'd-mmm',
			17 => 'mmm-yy',
			18 => 'h:mm AM/PM',
			19 => 'h:mm:ss AM/PM',
			20 => 'h:mm',
			21 => 'h:mm:ss',
			22 => 'm/d/yy h:mm',

			37 => '#,##0 ;(#,##0)',
			38 => '#,##0 ;[Red](#,##0)',
			39 => '#,##0.00;(#,##0.00)',
			40 => '#,##0.00;[Red](#,##0.00)',

			45 => 'mm:ss',
			46 => '[h]:mm:ss',
			47 => 'mmss.0',
			48 => '##0.0E+0',
			49 => '@',

			// CHT & CHS
			27 => '[$-404]e/m/d',
			30 => 'm/d/yy',
			36 => '[$-404]e/m/d',
			50 => '[$-404]e/m/d',
			57 => '[$-404]e/m/d',

			// THA
			59 => 't0',
			60 => 't0.00',
			61 =>'t#,##0',
			62 => 't#,##0.00',
			67 => 't0%',
			68 => 't0.00%',
			69 => 't# ?/?',
			70 => 't# ??/??'
		);
		private $Formats = array();

		private static $DateReplacements = array(
			'All' => array(
				'\\' => '',
				'am/pm' => 'A',
				'yyyy' => 'Y',
				'yy' => 'y',
				'mmmmm' => 'M',
				'mmmm' => 'F',
				'mmm' => 'M',
				':mm' => ':i',
				'mm' => 'm',
				'm' => 'n',
				'dddd' => 'l',
				'ddd' => 'D',
				'dd' => 'd',
				'd' => 'j',
				'ss' => 's',
				'.s' => ''
			),
			'24H' => array(
				'hh' => 'H',
				'h' => 'G'
			),
			'12H' => array(
				'hh' => 'h',
				'h' => 'G'
			)
		);

		private static $BaseDate = false;
		private static $DecimalSeparator = '.';
		private static $ThousandSeparator = '';
		private static $CurrencyCode = '';

		/**
		 * @var array Cache for already processed format strings
		 */
		private $ParsedFormatCache = array();

		/**
		 * @param string Path to file
		 * @param array Options:
		 *	TempDir => string Temporary directory path
		 *	ReturnDateTimeObjects => bool True => dates and times will be returned as PHP DateTime objects, false => as strings
		 */
		public function __construct($Filepath, array $Options = null)
		{
			if (!is_readable($Filepath))
			{
				throw new Exception('SpreadsheetReader_XLSX: File not readable ('.$Filepath.')');
			}

			$this -> TempDir = isset($Options['TempDir']) && is_writable($Options['TempDir']) ?
				$Options['TempDir'] :
				sys_get_temp_dir();

			$this -> TempDir = rtrim($this -> TempDir, DIRECTORY_SEPARATOR);
			$this -> TempDir = $this -> TempDir.DIRECTORY_SEPARATOR.uniqid().DIRECTORY_SEPARATOR;

			$Zip = new ZipArchive;
			$Status = $Zip -> open($Filepath);

			if ($Status !== true)
			{
				throw new Exception('SpreadsheetReader_XLSX: File not readable ('.$Filepath.') (Error '.$Status.')');
			}

			// Getting the general workbook information
			if ($Zip -> locateName('xl/workbook.xml') !== false)
			{
				$this -> WorkbookXML = new SimpleXMLElement($Zip -> getFromName('xl/workbook.xml'));
			}

			// Extracting the shared strings from the XLSX zip file
            if ($Zip -> locateName('xl/sharedStrings.xml', ZipArchive::FL_NOCASE) !== false) {
                $Zip -> extractTo($this -> TempDir, 'xl/sharedStrings.xml');
                $this -> TempFiles['sharedStrings'] = $this->TempDir.'xl'.DIRECTORY_SEPARATOR.'sharedStrings.xml';

				if (is_readable($this -> TempFiles['sharedStrings']))
				{
					$this -> PrepareSharedStringCache();
				}
			}

			$Sheets = $this -> Sheets();

			foreach ($this -> Sheets as $Index => $Name)
			{
				if ($Zip -> locateName('xl/worksheets/sheet'.$Index.'.xml') !== false)
				{
					$Zip -> extractTo($this -> TempDir, 'xl/worksheets/sheet'.$Index.'.xml');
					$this -> TempFiles['sheet' . $Index] = $this -> TempDir.'xl'.DIRECTORY_SEPARATOR.'worksheets'.DIRECTORY_SEPARATOR.'sheet'.$Index.'.xml';
				}
			}

			$this -> ChangeSheet(0);

			// If worksheet is present and is OK, parse the styles already
			if ($Zip -> locateName('xl/styles.xml') !== false)
			{
				$this -> StylesXML = new SimpleXMLElement($Zip -> getFromName('xl/styles.xml'));
				if ($this -> StylesXML && $this -> StylesXML -> cellXfs && $this -> StylesXML -> cellXfs -> xf)
				{
					foreach ($this -> StylesXML -> cellXfs -> xf as $Index => $XF)
					{
						// Format #0 is a special case - it is the "General" format that is applied regardless of applyNumberFormat
						if ($XF -> attributes() -> applyNumberFormat || (0 == (int)$XF -> attributes() -> numFmtId))
						{
							$FormatId = (int)$XF -> attributes() -> numFmtId;
							// If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
							$this -> Styles[] = $FormatId;
						}
						else
						{
							// 0 for "General" format
							$this -> Styles[] = 0;
						}
					}
				}
				
				if ($this -> StylesXML -> numFmts && $this -> StylesXML -> numFmts -> numFmt)
				{
					foreach ($this -> StylesXML -> numFmts -> numFmt as $Index => $NumFmt)
					{
						$this -> Formats[(int)$NumFmt -> attributes() -> numFmtId] = (string)$NumFmt -> attributes() -> formatCode;
					}
				}

				unset($this -> StylesXML);
			}

			$Zip -> close();

			// Setting base date
			if (!self::$BaseDate)
			{
				self::$BaseDate = new DateTime;
				self::$BaseDate -> setTimezone(new DateTimeZone('UTC'));
				self::$BaseDate -> setDate(1900, 1, 0);
				self::$BaseDate -> setTime(0, 0, 0);
			}

			// Decimal and thousand separators
			if (!self::$DecimalSeparator && !self::$ThousandSeparator && !self::$CurrencyCode)
			{
				$Locale = localeconv();
				self::$DecimalSeparator = $Locale['decimal_point'];
				self::$ThousandSeparator = $Locale['thousands_sep'];
				self::$CurrencyCode = $Locale['int_curr_symbol'];
			}

			if (function_exists('gmp_gcd'))
			{
				self::$RuntimeInfo['GMPSupported'] = true;
			}
		}

		/**
		 * Destructor, destroys all that remains (closes and deletes temp files)
		 */
		public function __destruct()
		{
			foreach ($this -> TempFiles as $TempFile)
			{
				@unlink($TempFile);
			}

			// Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
			if (strlen($this -> TempDir) > 2)
			{
				@rmdir($this -> TempDir.'xl'.DIRECTORY_SEPARATOR.'worksheets');
				@rmdir($this -> TempDir.'xl');
				@rmdir($this -> TempDir);
			}

			if ($this -> Worksheet && $this -> Worksheet instanceof XMLReader)
			{
				$this -> Worksheet -> close();
				unset($this -> Worksheet);
			}
			unset($this -> WorksheetPath);

			foreach ($this -> SharedStringReaders as $Reader)
			{
				if ($Reader -> Handle instanceof XMLReader)
				{
					$Reader -> Handle -> close();
				}
			}
			unset($this -> SharedStringReaders);

			if (isset($this -> StylesXML))
			{
				unset($this -> StylesXML);
			}
			if ($this -> WorkbookXML)
			{
				unset($this -> WorkbookXML);
			}
		}

		/**
		 * Retrieves an array with information about sheets in the current file
		 *
		 * @return array List of sheets (key is sheet index, value is name)
		 */
		public function Sheets()
		{
			if ($this -> Sheets === false)
			{
				$this -> Sheets = array();
				foreach ($this -> WorkbookXML -> sheets -> sheet as $Index => $Sheet)
				{
					$Attributes = $Sheet -> attributes('r', true);
					foreach ($Attributes as $Name => $Value)
					{
						if ($Name == 'id')
						{
							$SheetID = (int)str_replace('rId', '', (string)$Value);
							break;
						}
					}

					$this -> Sheets[$SheetID] = (string)$Sheet['name'];
				}
				ksort($this -> Sheets);
			}
			return array_values($this -> Sheets);
		}

		/**
		 * Changes the current sheet in the file to another
		 *
		 * @param int Sheet index
		 *
		 * @return bool True if sheet was successfully changed, false otherwise.
		 */
		public function ChangeSheet($Index)
		{
			$RealSheetIndex = false;
			$Sheets = $this -> Sheets();
			if (isset($Sheets[$Index]))
			{
				$SheetIndexes = array_keys($this -> Sheets);
				$RealSheetIndex = $SheetIndexes[$Index];
			}

			if ($RealSheetIndex !== false && isset($this -> TempFiles['sheet' . $RealSheetIndex]) && is_readable($this -> TempFiles['sheet' . $RealSheetIndex]))
			{
				$this -> WorksheetPath = $this -> TempFiles['sheet' . $RealSheetIndex];

				$this -> rewind();
				return true;
			}

			return false;
		}

		/**
		 * Creating shared string cache and chunk files
		 */
		private function PrepareSharedStringCache()
		{
			$SharedStrings = new XMLReader;
            $SharedStrings -> open($this -> TempFiles['sharedStrings']);

            $UsedMemory = 0;
			$CacheIndex = 0;
			$CacheValue = '';

			$chunkFile = null;
            $RowCount = 0;
			$FileNr = 0;

			while ($SharedStrings -> read())
			{
                switch ($SharedStrings -> name)
				{
					case 'sst':
						$this -> SharedStringCount = $SharedStrings -> getAttribute('uniqueCount');
						break;
					case 'si':
						if ($SharedStrings -> nodeType == XMLReader::END_ELEMENT)
						{
                            if(! is_null(self::SHARED_STRING_CACHE_LIMIT))
                            {
								if($UsedMemory < self::SHARED_STRING_CACHE_LIMIT)
								{
									$ValueSize = mb_strlen($CacheValue, '8bit');

									if ($UsedMemory + $ValueSize <= self::SHARED_STRING_CACHE_LIMIT)
									{
										$UsedMemory += $ValueSize;
										$this -> SharedStringCache[$CacheIndex] = $CacheValue;
									}
								}
                            }
							else
							{
								$this -> SharedStringCache[$CacheIndex] = $CacheValue;
							}

							$CacheIndex++;
							$CacheValue = '';
						}
						break;
					case 't':
						if ($SharedStrings -> nodeType == XMLReader::END_ELEMENT)
						{
							break;
						}

						$CacheValue .= $SharedStrings -> readString();

						// If we will cache everything we wont need chunk files
						if(! is_null(self::SHARED_STRING_CACHE_LIMIT))
						{
							if(is_resource($chunkFile))
							{
								if($RowCount < self::SHARED_STRING_CHUNK_SIZE - 1)
								{
									$RowCount++;
								}
								else
								{
									fwrite($chunkFile, '</root>');
									fclose($chunkFile);

									$FileNr++;
									$RowCount = 0;
								}
							}
							
							if(! is_resource($chunkFile))
							{
								$this -> TempFiles['strings-chunk-' . $FileNr] = $filename = $this -> TempDir.'xl/strings-chunk-' . $FileNr . '.xml';
								$chunkFile = fopen($filename, "w");
								fwrite($chunkFile, '<?xml version="1.0"?><root>');
							}

							fwrite($chunkFile, '<v>' . $CacheValue . '</v>');
						}
						break;
				}
			}

			if(is_resource($chunkFile))
			{
				fwrite($chunkFile, '</root>');
				fclose($chunkFile);
			}

			$SharedStrings -> close();

			return true;
		}
		
        private function GetSharedStringReader($Index)
        {
            $FileNr = 0;

			while(($FileNr + 1) * self::SHARED_STRING_CHUNK_SIZE <= $Index)
			{
				$FileNr++;
			}

            if(! isset($this -> SharedStringReaders[$FileNr]))
            {
                $this -> SharedStringsReaders[$FileNr] = new stdClass();
                $this -> SharedStringsReaders[$FileNr] -> Path = $this -> TempFiles['strings-chunk-' . $FileNr];
                $this -> SharedStringsReaders[$FileNr] -> BaseIndex = $FileNr * self::SHARED_STRING_CHUNK_SIZE;
                $this -> SharedStringsReaders[$FileNr] -> StringIndex = 0;
                $this -> SharedStringsReaders[$FileNr] -> LastValue = null;
                $this -> SharedStringsReaders[$FileNr] -> Handle = new XMLReader;

                $this -> SharedStringsReaders[$FileNr] -> Handle -> open($this -> SharedStringsReaders[$FileNr] -> Path);
            }

            return $FileNr;
        }

		/**
		 * Retrieves a shared string value by its index
		 *
		 * @param int Shared string index
		 *
		 * @return string Value
		 */
		private function GetSharedString($Index)
		{
			if (isset($this -> SharedStringCache[$Index]))
			{
				return $this -> SharedStringCache[$Index];
			}

            // If index of the desired string is larger than possible, don't even bother.
			if ($this -> SharedStringCount && ($Index >= $this -> SharedStringCount))
			{
				return '';
			}

            $ReaderId = $this -> GetSharedStringReader($Index);
            $Reader = $this -> SharedStringsReaders[$ReaderId];

            // If an index with the same value as the last already fetched is requested
			// (any further traversing the tree would get us further away from the node)
			if (($Index == $Reader -> BaseIndex + $Reader -> StringIndex) && ($Reader -> LastValue !== null))
			{
				return $Reader -> LastValue;
			}

			// If the desired index is before the current, rewind the XML
			if ($Reader -> BaseIndex + $Reader -> StringIndex > $Index)
			{
				$Reader -> Handle -> close();
				$Reader -> Handle -> open($Reader -> Path);
				$Reader -> StringIndex = 0;
				$Reader -> LastValue = null;
			}

			$Value = '';

            while ($Reader -> Handle -> read())
			{
				switch ($Reader -> Handle -> name)
				{
					case 'v':
						if ($Reader -> Handle -> nodeType == XMLReader::END_ELEMENT)
						{
							$Reader -> StringIndex++;

							if($Reader -> BaseIndex + $Reader -> StringIndex > $Index)
							{
								break 2;
							}
							break;
						}

						if($Reader -> BaseIndex + $Reader -> StringIndex == $Index)
						{
							$Value .= $Reader -> Handle -> readString();
							$Reader -> LastValue = $Value;
						}
						break;
				}
			}

			return $Value;
		}

		/**
		 * Formats the value according to the index
		 *
		 * @param string Cell value
		 * @param int Format index
		 *
		 * @return string Formatted cell value
		 */
		private function FormatValue($Value, $Index)
		{
			if (!is_numeric($Value))
			{
				return $Value;
			}

			if (isset($this -> Styles[$Index]) && ($this -> Styles[$Index] !== false))
			{
				$Index = $this -> Styles[$Index];
			}
			else
			{
				return $Value;
			}

			// A special case for the "General" format
			if ($Index == 0)
			{
				return $this -> GeneralFormat($Value);
			}

			$Format = array();

			if (isset($this -> ParsedFormatCache[$Index]))
			{
				$Format = $this -> ParsedFormatCache[$Index];
			}

			if (!$Format)
			{
				$Format = array(
					'Code' => false,
					'Type' => false,
					'Scale' => 1,
					'Thousands' => false,
					'Currency' => false
				);

				if (isset(self::$BuiltinFormats[$Index]))
				{
					$Format['Code'] = self::$BuiltinFormats[$Index];
				}
				elseif (isset($this -> Formats[$Index]))
				{
					$Format['Code'] = $this -> Formats[$Index];
				}

				// Format code found, now parsing the format
				if ($Format['Code'])
				{
					$Sections = explode(';', $Format['Code']);
					$Format['Code'] = $Sections[0];
	
					switch (count($Sections))
					{
						case 2:
							if ($Value < 0)
							{
								$Format['Code'] = $Sections[1];
							}
							break;
						case 3:
						case 4:
							if ($Value < 0)
							{
								$Format['Code'] = $Sections[1];
							}
							elseif ($Value == 0)
							{
								$Format['Code'] = $Sections[2];
							}
							break;
					}
				}

				// Stripping colors
				$Format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $Format['Code']));

				// Percentages
				if (substr($Format['Code'], -1) == '%')
				{
					$Format['Type'] = 'Percentage';
				}
				elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $Format['Code']))
				{
					$Format['Type'] = 'DateTime';

					$Format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $Format['Code']));
					$Format['Code'] = strtolower($Format['Code']);

					$Format['Code'] = strtr($Format['Code'], self::$DateReplacements['All']);
					if (strpos($Format['Code'], 'A') === false)
					{
						$Format['Code'] = strtr($Format['Code'], self::$DateReplacements['24H']);
					}
					else
					{
						$Format['Code'] = strtr($Format['Code'], self::$DateReplacements['12H']);
					}
				}
				elseif ($Format['Code'] == '[$EUR ]#,##0.00_-')
				{
					$Format['Type'] = 'Euro';
				}
				else
				{
					// Removing skipped characters
					$Format['Code'] = preg_replace('{_.}', '', $Format['Code']);
					// Removing unnecessary escaping
					$Format['Code'] = preg_replace("{\\\\}", '', $Format['Code']);
					// Removing string quotes
					$Format['Code'] = str_replace(array('"', '*'), '', $Format['Code']);
					// Removing thousands separator
					if (strpos($Format['Code'], '0,0') !== false || strpos($Format['Code'], '#,#') !== false)
					{
						$Format['Thousands'] = true;
					}
					$Format['Code'] = str_replace(array('0,0', '#,#'), array('00', '##'), $Format['Code']);

					// Scaling (Commas indicate the power)
					$Scale = 1;
					$Matches = array();
					if (preg_match('{(0|#)(,+)}', $Format['Code'], $Matches))
					{
						$Scale = pow(1000, strlen($Matches[2]));
						// Removing the commas
						$Format['Code'] = preg_replace(array('{0,+}', '{#,+}'), array('0', '#'), $Format['Code']);
					}

					$Format['Scale'] = $Scale;

					if (preg_match('{#?.*\?\/\?}', $Format['Code']))
					{
						$Format['Type'] = 'Fraction';
					}
					else
					{
						$Format['Code'] = str_replace('#', '', $Format['Code']);

						$Matches = array();
						if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $Format['Code']), $Matches))
						{
							$Integer = $Matches[1];
							$DecimalPoint = $Matches[2];
							$Decimals = $Matches[3];

							$Format['MinWidth'] = strlen($Integer) + strlen($DecimalPoint) + strlen($Decimals);
							$Format['Decimals'] = $Decimals;
							$Format['Precision'] = strlen($Format['Decimals']);
							$Format['Pattern'] = '%0'.$Format['MinWidth'].'.'.$Format['Precision'].'f';
						}
					}

					$Matches = array();
					if (preg_match('{\[\$(.*)\]}u', $Format['Code'], $Matches))
					{
						$CurrFormat = $Matches[0];
						$CurrCode = $Matches[1];
						$CurrCode = explode('-', $CurrCode);
						if ($CurrCode)
						{
							$CurrCode = $CurrCode[0];
						}

						if (!$CurrCode)
						{
							$CurrCode = self::$CurrencyCode;
						}

						$Format['Currency'] = $CurrCode;
					}
					$Format['Code'] = trim($Format['Code']);
				}

				$this -> ParsedFormatCache[$Index] = $Format;
			}

			// Applying format to value
			if ($Format)
			{
    			if ($Format['Code'] == '@')
    			{
        			return (string)$Value;
    			}
				// Percentages
				elseif ($Format['Type'] == 'Percentage')
				{
					if ($Format['Code'] === '0%')
					{
						$Value = round(100 * $Value, 0).'%';
					}
					else
					{
						$Value = sprintf('%.2f%%', round(100 * $Value, 2));
					}
				}
				// Dates and times
				elseif ($Format['Type'] == 'DateTime')
				{
					$Days = (int)$Value;
					// Correcting for Feb 29, 1900
					if ($Days > 60)
					{
						$Days--;
					}

					// At this point time is a fraction of a day
					$Time = ($Value - (int)$Value);
					$Seconds = 0;
					if ($Time)
					{
						// Here time is converted to seconds
						// Some loss of precision will occur
						$Seconds = (int)($Time * 86400);
					}

					$Value = clone self::$BaseDate;
					$Value -> add(new DateInterval('P'.$Days.'D'.($Seconds ? 'T'.$Seconds.'S' : '')));

					if (!$this -> Options['ReturnDateTimeObjects'])
					{
						$Value = $Value -> format($Format['Code']);
					}
					else
					{
						// A DateTime object is returned
					}
				}
				elseif ($Format['Type'] == 'Euro')
				{
					$Value = 'EUR '.sprintf('%1.2f', $Value);
				}
				else
				{
					// Fractional numbers
					if ($Format['Type'] == 'Fraction' && ($Value != (int)$Value))
					{
						$Integer = floor(abs($Value));
						$Decimal = fmod(abs($Value), 1);
						// Removing the integer part and decimal point
						$Decimal *= pow(10, strlen($Decimal) - 2);
						$DecimalDivisor = pow(10, strlen($Decimal));

						if (self::$RuntimeInfo['GMPSupported'])
						{
							$GCD = gmp_strval(gmp_gcd($Decimal, $DecimalDivisor));
						}
						else
						{
							$GCD = self::GCD($Decimal, $DecimalDivisor);
						}

						$AdjDecimal = $DecimalPart/$GCD;
						$AdjDecimalDivisor = $DecimalDivisor/$GCD;

						if (
							strpos($Format['Code'], '0') !== false || 
							strpos($Format['Code'], '#') !== false ||
							substr($Format['Code'], 0, 3) == '? ?'
						)
						{
							// The integer part is shown separately apart from the fraction
							$Value = ($Value < 0 ? '-' : '').
								$Integer ? $Integer.' ' : ''.
								$AdjDecimal.'/'.
								$AdjDecimalDivisor;
						}
						else
						{
							// The fraction includes the integer part
							$AdjDecimal += $Integer * $AdjDecimalDivisor;
							$Value = ($Value < 0 ? '-' : '').
								$AdjDecimal.'/'.
								$AdjDecimalDivisor;
						}
					}
					else
					{
						// Scaling
						$Value = $Value / $Format['Scale'];

						if (!empty($Format['MinWidth']) && $Format['Decimals'])
						{
							if ($Format['Thousands'])
							{
								$Value = number_format($Value, $Format['Precision'],
									self::$DecimalSeparator, self::$ThousandSeparator);
							}
							else
							{
								$Value = sprintf($Format['Pattern'], $Value);
							}

							$Value = preg_replace('{(0+)(\.?)(0*)}', $Value, $Format['Code']);
						}
					}

					// Currency/Accounting
					if ($Format['Currency'])
					{
						$Value = preg_replace('', $Format['Currency'], $Value);
					}
				}
				
			}

			return $Value;
		}

		/**
		 * Attempts to approximate Excel's "general" format.
		 *
		 * @param mixed Value
		 *
		 * @return mixed Result
		 */
		public function GeneralFormat($Value)
		{
			// Numeric format
			if (is_numeric($Value))
			{
				$Value = (float)$Value;
			}
			return $Value;
		}

		// !Iterator interface methods
		/** 
		 * Rewind the Iterator to the first element.
		 * Similar to the reset() function for arrays in PHP
		 */ 
		public function rewind()
		{
			// Removed the check whether $this -> Index == 0 otherwise ChangeSheet doesn't work properly

			// If the worksheet was already iterated, XML file is reopened.
			// Otherwise it should be at the beginning anyway
			if ($this -> Worksheet instanceof XMLReader)
			{
				$this -> Worksheet -> close();
			}
			else
			{
				$this -> Worksheet = new XMLReader;
			}

			$this -> Worksheet -> open($this -> WorksheetPath);

			$this -> Valid = true;
			$this -> RowOpen = false;
			$this -> CurrentRow = false;
			$this -> Index = 0;
		}

		/**
		 * Return the current element.
		 * Similar to the current() function for arrays in PHP
		 *
		 * @return mixed current element from the collection
		 */
		public function current()
		{
			if ($this -> Index == 0 && $this -> CurrentRow === false)
			{
				$this -> next();
				$this -> Index--;
			}
			return $this -> CurrentRow;
		}

		/** 
		 * Move forward to next element. 
		 * Similar to the next() function for arrays in PHP 
		 */ 
		public function next()
		{
			$this -> Index++;

			$this -> CurrentRow = array();

			if (!$this -> RowOpen)
			{
				while ($this -> Valid = $this -> Worksheet -> read())
				{
					if ($this -> Worksheet -> name == 'row')
					{
						// Getting the row spanning area (stored as e.g., 1:12)
						// so that the last cells will be present, even if empty
						$RowSpans = $this -> Worksheet -> getAttribute('spans');
						if ($RowSpans)
						{
							$RowSpans = explode(':', $RowSpans);
							$CurrentRowColumnCount = $RowSpans[1];
						}
						else
						{
							$CurrentRowColumnCount = 0;
						}

						if ($CurrentRowColumnCount > 0)
						{
							$this -> CurrentRow = array_fill(0, $CurrentRowColumnCount, '');
						}

						$this -> RowOpen = true;
						break;
					}
				}
			}

			// Reading the necessary row, if found
			if ($this -> RowOpen)
			{
				// These two are needed to control for empty cells
				$MaxIndex = 0;
				$CellCount = 0;

				$CellHasSharedString = false;

				while ($this -> Valid = $this -> Worksheet -> read())
				{
					switch ($this -> Worksheet -> name)
					{
						// End of row
						case 'row':
							if ($this -> Worksheet -> nodeType == XMLReader::END_ELEMENT)
							{
								$this -> RowOpen = false;
								break 2;
							}
							break;
						// Cell
						case 'c':
							// If it is a closing tag, skip it
							if ($this -> Worksheet -> nodeType == XMLReader::END_ELEMENT)
							{
								break;
							}

							$StyleId = (int)$this -> Worksheet -> getAttribute('s');

							// Get the index of the cell
							$Index = $this -> Worksheet -> getAttribute('r');
							$Letter = preg_replace('{[^[:alpha:]]}S', '', $Index);
							$Index = self::IndexFromColumnLetter($Letter);

							// Determine cell type
							if ($this -> Worksheet -> getAttribute('t') == self::CELL_TYPE_SHARED_STR)
							{
								$CellHasSharedString = true;
							}
							else
							{
								$CellHasSharedString = false;
							}

							$this -> CurrentRow[$Index] = '';

							$CellCount++;
							if ($Index > $MaxIndex)
							{
								$MaxIndex = $Index;
							}

							break;
						// Cell value
						case 'v':
						case 'is':
							if ($this -> Worksheet -> nodeType == XMLReader::END_ELEMENT)
							{
								break;
							}

							$Value = $this -> Worksheet -> readString();

							if ($CellHasSharedString)
							{
								$Value = $this -> GetSharedString($Value);
							}

							// Format value if necessary
							if ($Value !== '' && $StyleId && isset($this -> Styles[$StyleId]))
							{
								$Value = $this -> FormatValue($Value, $StyleId);
							}
							elseif ($Value)
							{
								$Value = $this -> GeneralFormat($Value);
							}

							$this -> CurrentRow[$Index] = $Value;
							break;
					}
				}

				// Adding empty cells, if necessary
				// Only empty cells inbetween and on the left side are added
				if ($MaxIndex + 1 > $CellCount)
				{
					$this -> CurrentRow = $this -> CurrentRow + array_fill(0, $MaxIndex + 1, '');
					ksort($this -> CurrentRow);
				}
			}

			return $this -> CurrentRow;
		}

		/** 
		 * Return the identifying key of the current element.
		 * Similar to the key() function for arrays in PHP
		 *
		 * @return mixed either an integer or a string
		 */ 
		public function key()
		{
			return $this -> Index;
		}

		/** 
		 * Check if there is a current element after calls to rewind() or next().
		 * Used to check if we've iterated to the end of the collection
		 *
		 * @return boolean FALSE if there's nothing more to iterate over
		 */ 
		public function valid()
		{
			return $this -> Valid;
		}

		// !Countable interface method
		/**
		 * Ostensibly should return the count of the contained items but this just returns the number
		 * of rows read so far. It's not really correct but at least coherent.
		 */
		public function count()
		{
			return $this -> Index + 1;
		}

		/**
		 * Takes the column letter and converts it to a numerical index (0-based)
		 *
		 * @param string Letter(s) to convert
		 *
		 * @return mixed Numeric index (0-based) or boolean false if it cannot be calculated
		 */
		public static function IndexFromColumnLetter($Letter)
		{
			$Powers = array();

			$Letter = strtoupper($Letter);

			$Result = 0;
			for ($i = strlen($Letter) - 1, $j = 0; $i >= 0; $i--, $j++)
			{
				$Ord = ord($Letter[$i]) - 64;
				if ($Ord > 26)
				{
					// Something is very, very wrong
					return false;
				}
				$Result += $Ord * pow(26, $j);
			}
			return $Result - 1;
		}

		/**
		 * Helper function for greatest common divisor calculation in case GMP extension is
		 *	not enabled
		 *
		 * @param int Number #1
		 * @param int Number #2
		 *
		 * @param int Greatest common divisor
		 */
		public static function GCD($A, $B)
		{
			$A = abs($A);
			$B = abs($B);
			if ($A + $B == 0)
			{
				return 0;
			}
			else
			{
				$C = 1;

				while ($A > 0)
				{
					$C = $A;
					$A = $B % $A;
					$B = $C;
				}

				return $C;
			}
		}
	}
?>
