use strict;
use warnings;
use Essent;

my $dir = "c:\\Dokumente und Einstellungen\\huesemann.POLYINTERN\\Eigene Dateien\\workspace\\excelplay\\";
my $todo = $dir."TODO.txt";
my @ignore = qw(parse_TODO.pl);
#my @ignore = qw();


my @perls;
@perls = File::get_by_ext($dir, 'pl');
push @perls, File::get_by_ext($dir, 'pm');

### parse each file
my @parse_file;
my $file_to_sign = 0;
my @parse_results;

# TODO add line-number
# no file_signed, if no TODO at all

foreach my $file (@perls) {
	next if grep {$file =~ /$_/ } @ignore;
	my @filecontent = File::readfile($file);
	foreach my $line (@filecontent) {
		if ($line =~ /(TODO.*)/) {
			push @parse_file, Data::remove_ws $1;
			$file_to_sign = 1;
		}
	}
	if ($file_to_sign) {
		unshift @parse_file, "===========";
		unshift @parse_file, $file;
		# parsed TODOs
		push @parse_file, "===========\n";
		$file_to_sign = 0;
	}
	push @parse_results, @parse_file;
	@parse_file = ();
}

@parse_results = map {$_."\n"} (@parse_results);
unshift @parse_results, Process::getTime()."\n";

#my $todo_string = join "\n", @parse_results;

File::writefile($todo, @parse_results);
