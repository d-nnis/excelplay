use strict;
use warnings;
use Essent;

#my $dir = "c:\\Dokumente und Einstellungen\\huesemann.POLYINTERN\\Eigene Dateien\\workspace\\excelplay\\";
#my $dir = "f:\\Users\\d-nnis\\workspace\\excelplay\\";
my $dir = ".\\";
my $todo = $dir."TODO.txt";
my @ignore = qw(parse_TODO.pl);
my $delete_tmp = 1;
# präzise genug
# Ende von Package und sub nicht erkannt




my @perls;
@perls = File::get_by_ext($dir, 'pl');
push @perls, File::get_by_ext($dir, 'pm');

### parse each file
my @parse_file;
my $file_to_sign = 0;
my $follow_line = 0;
my @parse_results;
my $line_number;
my $package = '';
my $sub = '';
my %context = ();
my $context_sign = '';

foreach my $file (@perls) {
	next if grep {$file =~ /$_/ } @ignore;
	my @filecontent = File::readfile($file);
	$line_number = 0;
	$package = '';
	$sub = '';
	foreach my $line (@filecontent) {
		$line_number++;
		$package = ("_$line_number: ".$1) if $line =~ /\s*(package\s\w+)\s*;$/;
		$sub = ("_$line_number: ".$1) if $line =~ /\s*(sub\s\w+)\s*{$/;
		$context_sign = $package."-".$sub;
		
		unless ($context{$context_sign})  {
			$context{$context_sign} = 0;
		}
		if ($line =~ /(TODO.*)/) {
			push @parse_file, $context_sign if $context{$context_sign} == 0 && length $context_sign > 1;
			$context{$context_sign} = 1;
			push @parse_file, "$line_number: ".Data::remove_ws $1;
			$file_to_sign = 1;
			$follow_line = 1;
		} elsif ($line =~ /\s*#+\s\w/ && $follow_line) {
			# push if '   # irgendwas'
			push @parse_file, "$line_number: ".Data::remove_ws $line;
		} else {
			$follow_line = 0;
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


if ($delete_tmp) {
    my @dels = File::get_by_ext($dir, 'html');
    foreach (@dels) {
        if ($_ =~ /^(tmp.*\.html)/) {
            push @parse_results, "DELETING $dir$1";
            print "DELETING $dir$1\n";
            unlink $dir.$1;    
        }
    } 
}

@parse_results = map {$_."\n"} (@parse_results);
unshift @parse_results, Process::getTime()."\n";

File::writefile($todo, @parse_results);

print "ende\n";
Process::confirm();