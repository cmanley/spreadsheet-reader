<?php
if (isset($argv)) {
	error_log("Running outside of phpunit. Consider using phpunit.\n");
	class PHPUnit_Framework_TestCase {}
}

$terminal_encoding = mb_internal_encoding();
if (isset($_SERVER['LANG'])) {	# e.g.: LANG=en_US.UTF-8
	$terminal_encoding = preg_replace('/^[^\.]+\.(.+)$/', '$1', $_SERVER['LANG']);
	mb_internal_encoding($terminal_encoding);
	ini_set('default_charset', mb_internal_encoding());
}


class Test extends PHPUnit_Framework_TestCase
{
	const CLASS_NAME = 'SpreadsheetReader';
	const FILE = __DIR__ . '/../SpreadsheetReader.php';
	public static $files = array(	# These are a copy of some free Microsoft sample online
		'users.csv',	# Saved in UTF-8 encoding.
		'users.ods',	# FIXME: too many trailing blank lines and columns are read with this format.
		'users.xls',	# Always 'ANSI' 8 bit encoding.
		'users.xlsx',
		'Sheets.xlsx',	# The sheet file name starts with a capitical S.
		'inconsistant_case.xlsx',	# The actual sheet file name is Sheet1.xml but the rels file refers to it as 'sheet1.xml'. Crystal Reports does this.
		'ods.zip',
		'xlsx.zip',
	);

	public function testRequire() {
		$this->assertFileExists(static::FILE);
		include(static::FILE);
		$class = static::CLASS_NAME;
		$this->assertTrue(class_exists($class), 'Check that class name "' . $class . '" exists after include.');
	}

	public function testImplementsInterface() {
		$class = static::CLASS_NAME;
		$interfaces = array(
			'SeekableIterator',
			'Countable',
		);
		foreach ($interfaces as $interface) {
			$this->assertTrue(is_subclass_of ($class, '\\' . $interface, true), 'Check that class "' . $class . '" implements \\' . $interface . '.');
		}
	}

	public function testMethods() {
		$class = static::CLASS_NAME;
		$methods = array(
			'__construct'		=> ReflectionMethod::IS_PUBLIC,
			'getReaderType'		=> ReflectionMethod::IS_PUBLIC,
			'Sheets'			=> ReflectionMethod::IS_PUBLIC,
			'ChangeSheet'		=> ReflectionMethod::IS_PUBLIC,
			'Load'				=> ReflectionMethod::IS_PRIVATE | ReflectionMethod::IS_STATIC,
			'rewind'			=> ReflectionMethod::IS_PUBLIC,
			'current'			=> ReflectionMethod::IS_PUBLIC,
			'next'				=> ReflectionMethod::IS_PUBLIC,
			'key'				=> ReflectionMethod::IS_PUBLIC,
			'valid'				=> ReflectionMethod::IS_PUBLIC,
			'count'				=> ReflectionMethod::IS_PUBLIC,
			'seek'				=> ReflectionMethod::IS_PUBLIC,
		);
		foreach ($methods as $name => $expected_modifiers) {
			$exists = method_exists($class, $name);
			$this->assertTrue($exists, "Check method $class::$name() exists.");
			if ($exists) {
				$method = new ReflectionMethod($class, $name);
				$actual_modifiers = $method->getModifiers() & (
					ReflectionMethod::IS_STATIC |
					ReflectionMethod::IS_PUBLIC |
					ReflectionMethod::IS_PROTECTED |
					ReflectionMethod::IS_PRIVATE |
					ReflectionMethod::IS_ABSTRACT |
					ReflectionMethod::IS_FINAL
				);
				#error_log("$name expected: " . $expected_modifiers);
				#error_log("$name actual:   " . $actual_modifiers);
				$this->assertEquals($expected_modifiers, $actual_modifiers, "Expected $class::$name() modifiers to be \"" . join(' ', Reflection::getModifierNames($expected_modifiers)) . '" but got "' . join(' ', Reflection::getModifierNames($actual_modifiers)) . '" instead.');
			}
		}
	}

	public function testIterate() {
		$class = static::CLASS_NAME;
		$script_encoding = 'UTF-8';
		$expect_first_row_first_col = 'First Name';
		$expect_first_row_last_col = 'Time';
		$expect_last_row_first_col = 'Frédéric';
		#$expect_last_row_last_col = '05:05:59';	# Times and dates are inconsistently formatted.
		$expect_col_count = 16;
		$expect_row_count = 6;	# count() is wrong for all types except XLS by design.
		foreach (static::$files as $file) {
			$reader = new $class("data/$file");
			$type = $reader->getReaderType();
			$file_encoding = $type == $class::TYPE_XLS ? 'Windows-1252' : 'UTF-8';	# XLS stores data in 8 bit Windows 'ANSI' format, but which exact encoding is unknown. All other files are in UTF-8.

			$row = $reader->current();

			# FIXME: ODS returns 256 columns.
			($type == $class::TYPE_ODS) && $row = array_splice($row, 0, $expect_col_count);

			$this->assertEquals($expect_col_count, count($row), "Expected $expect_col_count columns in $file but got " . count($row) . ' instead');

			$got = mb_convert_encoding(reset($row), $script_encoding, $file_encoding);
			$this->assertEquals($expect_first_row_first_col, $got, "First record of $file contains \"$expect_first_row_first_col\"");

			$got = mb_convert_encoding(end($row), $script_encoding, $file_encoding);
			$this->assertEquals($expect_first_row_last_col, $got, "First record of $file contains \"$expect_first_row_last_col\"");

			# Test count() before iteration.
			($type == $class::TYPE_XLS) && $this->assertEquals($expect_row_count, $reader->count(), "Expected $expect_row_count rows in $file but got " . $reader->count() . ' instead');

			# Iterate to end.
			$y = 0;
			foreach($reader as $row) {
				$y++;
				if (($type == $class::TYPE_ODS) && ($y >= $expect_row_count)) { # FIXME: ODS reads blank rows past end of data.
					break;
				}
			}

			$got = mb_convert_encoding(reset($row), $script_encoding, $file_encoding);
			$this->assertEquals($expect_last_row_first_col, $got, "Last record of $file contains \"$expect_last_row_first_col\"");
			#$got = mb_convert_encoding(end($row), $script_encoding, $file_encoding);
			#$this->assertEquals($expect_last_row_last_col, $got, "Last record of $file contains \"$expect_last_row_last_col\"");

			# Test rewind()
			$reader->rewind();
			$row = $reader->current();
			$got = mb_convert_encoding(reset($row), $script_encoding, $file_encoding);
			$this->assertEquals($expect_first_row_first_col, $got, "First record of $file contains \"$expect_first_row_first_col\" after rewind.");
		}
	}

}


if (isset($argv)) {
	require_once(Test::FILE);
	$class = Test::CLASS_NAME;
	if (count($argv) >= 2) {
		$file = $argv[1];
		$reader = new $class($file);
		$type = $reader->getReaderType();
		$file_encoding = ($type == $class::TYPE_XLS) ? 'Windows-1252' : 'UTF-8';	# XLS stores data in 8bit ANSI format.
		$y = 0;
		$max_rows = 1;
		foreach($reader as $row) {
			print mb_convert_encoding(print_r($row,1), $terminal_encoding, $file_encoding);
			if ($max_rows && (++$y >= $max_rows)) {
				break;
			}
		}
	}
	else {
		$files = Test::$files;
		foreach ($files as $file) {
			print $file . ":\n";
			$reader = new $class("data/$file");
			$type = $reader->getReaderType();
			print "type: $type\n";
			$file_encoding = ($type == $class::TYPE_XLS) ? 'Windows-1252' : 'UTF-8';
			#$row = $reader->current();
			#print_r($row);
			foreach($reader as $row) {
				print mb_convert_encoding(print_r($row,1), $terminal_encoding, $file_encoding);
			}
			print "\n\n";
		}
	}
}
