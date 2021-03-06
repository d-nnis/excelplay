use feature qw(switch);
use strict;
use warnings;

#use IO::Handle;

print "Modul essent.pm\n";

{
	package File;
	
	sub readfile {
		my $rfile = $_[0];
		my $csv = 0;
		my $config = 0;
		my $forms_job = 0;
		if (defined $_[1]) {
			$csv = 1 if $_[1] eq 'csv';
			$config = 1 if $_[1] eq 'config';
			$forms_job = 1 if $_[1] eq 'forms_job';
		}
		my $call_from = Data::extractfilename($0);
		$call_from = $_[2] if defined $_[2];
		if (!open (RFILE, '<', $rfile) ) {
			warn "\n!!! Achtung: Kann $rfile nicht oeffnen: $!\n";
			while (<STDIN> eq '') {}
			readfile($_);
		}
		print "Datei oeffnen: $rfile\n";
		my @readarray;
		my %config;
		my %forms_job;
		my $forms_part;
		if (-e $rfile) {
			my $i = 0;
			my $read_config_part = 0;
			my $read_job_forms = 0;
			while (<RFILE>) {
				my $fline = $_;
				if ($csv) {
					# Kommentarzeilen mit '#' herausfiltern
					push (@readarray, $fline) if ($fline !~ /^#.*/);
					$i++;
				} elsif ($config) {
					if ($fline !~ /^#.*/ && $fline =~ /^\S.*/) {
						$read_config_part = 0 if ($fline =~ /\[/);
						if ($read_config_part) {
							(my $key, my $value) = (split(/;/, $fline))[0,1];
							# Pfad parsen: normaler Pfad zu Perl-lesbaren-Format - noetig?
							#$value =~ s/\\/\\\\/g;
							$value = Data::path_perl($value);
							$config{$key} = Data::remove_ws($value);
						}
						$read_config_part = 1 if ($fline =~ /\[$call_from\]/);
						$read_config_part = 1 if ($fline =~ /\[allgemein\]/);
					}
				} elsif ($forms_job) {
					unless ($fline =~ /^#/) {
						if ($fline =~ /(\[.*\])/) {
							$forms_part = $1;
						} else {
							push @{$forms_job{$forms_part}}, $fline;
						}
					}
				} else {
					push (@readarray,$fline);
					#$readarray[$i]=$fline;
					#$i++;
				}
			}
			close RFILE;
		} else {
			print "\n Datei $rfile existiert nicht!! \n";
		}
		if ($config) {
			return %config;
		} elsif ($forms_job) {
			return %forms_job;
		} else {
			return @readarray;
		}
	}
	
	
	sub writeSPSfile {
		my $file = shift;
		my @lines = map{' '.$_} @_;
		my $sps_project = Data::extractfilename($file);
		my $win_file = Data::path_win($file);
		my @sps_header = ("TITLE '$sps_project'.\n",
			"COMMENT 'openTRS conversion'.\n",
			"DATA LIST FIXED FILE = '$win_file' RECORDS=1\n",
			" /1\n");
		unshift (@lines, @sps_header);
		# sps-foot
		my @sps_foot = (".\n", "EXECUTE.\n");
		push (@lines, @sps_foot);
		writefile($file, @lines);
	}
	
	sub writefile_exist {
		my $file = shift;
		my @lines = @_;
		if (-e $file) {
			print "!!! Datei existiert: ", $file, "\nNeuer Schreibversuch mit Tastendruck\n";
			while (<STDIN> eq '') {}
			writefile_exist($file, @lines);
		}
		writefile($file, @lines);
	}
	
	sub writefile {
		my $file = shift;
		#my @lines = @_[1 .. $#_];	# alle Elemente 1 bis Ende
		my @lines = @_;	# alle Elemente 0 bis Ende
		print "write file ", $file;
		if (!open (WFILE, '>', $file) ) {
			print "\n!!! Achtung: Kann $file nicht oeffnen: $!\nNeuer Versuch Tastendruck";
			while (<STDIN> eq '') {}
			writefile($file, @lines);
		}
		open (WFILE, '>', $file);
		print WFILE @lines;
        my $line_num = 0;
        if (scalar @lines > 1) {
            $line_num = scalar @lines;
        } else {
            foreach (@lines) {
				my @arr = split /\n/, $_;
                my $scal = scalar @arr;
                $line_num += $scal if $line_num < $scal;
            }            
        }
		print " lines: ", $line_num, ".\n";
		close WFILE;
	}
	
	sub writefile_count {
		my $file = shift;
		my @lines = @_;
		my $file_count = 0;
		my $ext = $file;
		while ($ext =~ /\.(.+)/) {
			$ext = $1;
		}
		my $file_stem = substr $file, 0, (length $ext) * -1 - 1;
		while ( -e $file) {
			$file = $file_stem . $file_count . "." . $ext;
			$file_count++;
		}
		File::writefile($file, @lines);
	}
	
	sub writefile_app {
		my $file = $_[0];
		my @lines = @_[1 .. $#_];	# alle Elemente 1 bis Ende
		print "write to file  ", $file;
		if (!open (WFILE, '>>', $file) ) {
		print "\n!!! Achtung: Kann $file nicht oeffnen: $!\n";
			while (<STDIN> eq '') {}
			writefile(@lines);
		}
		
		print WFILE @lines;
		print " lines added: ", scalar @lines, ".\n";
		close WFILE;
	}
	
	sub get_by_ext {
		my $dir = shift;
		my $ext = shift;
		# qr!
		my $regex = join('', '(.*\.', $ext, '$)');
		opendir(DIR, $dir);
		my @files = readdir(DIR);
		my @erg;
		foreach my $file (@files) {
			my $match = $file =~ /$regex/i;
			push (@erg, $1) if ($match);
		}
		# sortiere in alpahbetische Ordnung
		@erg = sort { lc($a) cmp lc($b)} @erg;
		closedir(DIR);
		return @erg;
	}
	
	## get_subdirs
	## out: subdirs (relative path)
	##  TODO optionale absolute path
	##  option: depth of subdir-search, depths-first
	sub get_subdirs {
		my $dir = shift;
		my $option = $_[0] if defined $_[0];
		opendir(DIR, $dir);
		my @content = readdir(DIR);
		my @subdirs;
		foreach my $cont (@content) {
			if (-d $dir.$cont && $cont !~ /^\./) {
				given ($option) {
					when ('num') { push @subdirs, $cont if $cont =~ /^\d/; }
					default { push @subdirs, $cont; }
				}
			}
		}
		@subdirs = sort { lc($a) cmp lc($b)} @subdirs;
		return @subdirs;
	}

	## get_subdirs
	## out: subdirs (relative path)
	##  TODO optionale absolute path
	##  option: depth of subdir-search, depths-first
	## default: depth=0 - no subdirs of subdir
	## TODO depth>0 is too slow!
	sub get_subdirs2 {
		my $dir = shift;
		my $depth = shift;
		$depth = 0 unless defined $depth;
		my $option;	# TODO: options-handling
		$dir .= "\\" unless $dir =~ /\\$/;
		opendir(DIR, $dir);
		my @content = readdir(DIR);
		my @subdirs;
		foreach my $cont (@content) {
			if (-d $dir.$cont && $cont !~ /^\./) {
				push @subdirs, $dir.$cont;
				push @subdirs, get_subdirs2($dir.$cont, $depth-1) if $depth > 0;
				# if ($option eq 'num') { push @subdirs, $dir.$cont if $cont =~ /^\d/; }
			}
			
		}
		@subdirs = sort { lc($a) cmp lc($b)} @subdirs;
		return @subdirs;
	}
	
	sub one {
		print "one\n";
		print $_[0], "\n";
		return join ('', $_[0], "_one_");
	}
}

{
	package Process;
	
	sub confirm {
		my $eingabe='';
		until ( $eingabe eq 'j' ) {
			print "Fortfahren mit 'j'..\n>";
			$eingabe = <STDIN>;
			chomp $eingabe;
		}
	}
	
	sub confirmJN {
		my $eingabe='';
		my @exp_keys = ('j','n');
		until (grep {$eingabe eq $_} @exp_keys) {
			print "(j)a oder (n)ein... \n>";
			$eingabe = <STDIN>;
			chomp $eingabe;
		}
		if ($eingabe eq 'j') {
			return 1;
		} else {
			return 0;
		}
	}
    
	sub confirm_numcount {
		my $max = shift;
		my $eingabe='';
		my @exp_keys = (1..$max);
		until (grep {$eingabe eq $_} @exp_keys) {
			print (join ",", @exp_keys);
			print "\n>";
			$eingabe = <STDIN>;
			chomp $eingabe;
		}
		return $eingabe+0;
	}
	

    sub confirmJNdie {
		my $eingabe='';
		my @exp_keys = ('j','n');
		until (grep {$eingabe eq $_} @exp_keys) {
			print "(j)a, (n)ein oder (c)Abbruch... \n>";
			$eingabe = <STDIN>;
			chomp $eingabe;
		}
        given ($eingabe) {
            when ('c') { die "Abbruch\n"; }
            when ('j') { return 1; }
            when ('n') { return 0; }
        }
    }
    
	sub getTime {
		my $sec; my $min; my $hour; my $mday; my $mon; my $year; my $wday; my $yday; my $isdst;
		my @abbr = qw( Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec );
		($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
		#return ($year +=1900), " ", $abbr[$mon], " ", $mday, ", ", $hour, ":", $min, ":", $sec;
		return ($year +=1900). " ". $abbr[$mon]. " ". $mday. ", ". $hour. ":". $min. ":". $sec;
	}
}

{
	package Data;
	
    
    sub parse_csv {
        my $dataref = shift;
        my $field_sep = shift;
        my $has_field_names = shift;
        my $value_enclosed = shift || "";
        my $chomp = shift || 1;
        
        my $sep = $value_enclosed . $field_sep . $value_enclosed;
        
        my @data = @{$dataref};
        
        my @field_names;
        if ($has_field_names) {
            my $headline = shift @data;
            #@field_names = split /$field_sep/, $headline;
            #my $sep = $value_enclosed . $field_sep . $value_enclosed;
            chomp $headline if $chomp;
            @field_names = split /$sep/, $headline;
            # funktioniert noch nicht mit $value_enclosed!
        }
        
        my @csv_matrix;
        foreach my $line (@data) {
            my @line_arr = split /$sep/, $line;
            chomp $line_arr[-1] if $chomp;
            push (@csv_matrix, [@line_arr]);
        }
        
        
        my %field;
        my $i = 0;
        foreach my $field (@field_names) {
            $field{$field} = $i++;
        }
        
        return \%field, \@csv_matrix;
        
    }
    
	# setzt neues Array aus den ~distinkten Elementen in original Reihenfolge
	sub distinct {
		my @arr = @_;
		my @distinct;
		push (@distinct, shift @arr);
		foreach (@arr) {
			push (@distinct, $_) unless $distinct[-1] eq $_;	
		}
		return @distinct;
	}
	
	# in: Array
	# out: Array mit Position ~distinktiver Elemente in input-Reihenfolge 
	sub distinct_pos {
		my @arr = @_;
		my $temp = shift @arr;
		my $size = scalar @arr;
		my @distinct_pos = (0);
		#push (@distinct, shift @arr);
		for (my $i = 0; $i < $size; $i++) {
			push (@distinct_pos, $i+1) unless $arr[0] eq $temp;
			$temp = shift @arr; 
			}
		return @distinct_pos;
	}
	# my @elems = qw (1 1 2 3 3 3 3 4 4 4 3);
	
	
	# entferne preceding und trailing whitespaces (white space (new line, carriage return, space, tab, form feed))
	sub remove_ws {
		my $string = '';
		if (defined $_[0]) {
			$string = $_[0];
			$string =~ s/^\s+//;
			$string =~ s/\s+$//;
		}
		return $string;
	}
	
	sub remove_ws_arr {
		my @in = @_;
		my @out = ();
		foreach (@in) {
			my $string = $_;
			$string =~ s/^\s+//;
			$string =~ s/\s+$//;
			push @out, $string;
		}
		return @out;
	}
		
	sub chomping {
		chomp @_;
		return @_;
	}
	
	# stamm ohne erw
	sub get_fname {
		return $1 if ($_[0] =~ /(.+)\./);
		warn "!! get_fname-error\n";
	}
	
	sub get_extname {
		#return $1 if ($_[0] =~ /\.(\w{3})/);
		return $1 if ($_[0] =~ /\.(\w+)$/);
		warn "!! get_extname-error\n";
	}
	
	# extrahier Dateinamen aus vollem Pfad:
	#  Eintrag nach letztem '\', bzw.
	#  kompletten Eintrag, wenn kein '\'
	#  Entferne Whitespaces
	sub extractfilename {
		my $filepath = $_[0];
		while ($filepath =~ /\\(.*)/) {
			$filepath = $1;
		}
		return remove_ws($filepath);
	}
	
	sub extractPath {
		my $fullpath = shift;
		# greedy: nimmt alles bis zum __letzten__ Backslash!
		(my $path) = $fullpath =~ /(.*\\)/;
		#while ($fullpath =~ /(.*\\)/) {
		#	$path = $path . $1;
		#	$fullpath =~ /.+\\(.+)/;
		#	$fullpath = $1;
		#}
		return $path;
		
	}
	
	# suche letztes dir vor filename raus
	sub extract_subdir_filename {
		my $filepath = $_[0];
		while ($filepath =~ /\\(.*\\.*)/) {
			$filepath = $1;
		}
		return remove_ws($filepath);
	}
	
	# SPSS-Format f�r substr umbauen:
	# '9-13 (A)'
	# (9,13)
	# (9,5)
	sub spss_position {
		my $val = shift;
		my $switch = shift;
		my $fi;
		my $se;
		my $th = '';
		given ($switch) {
			when (/format/) {
				# '1-10 (A)'
				($fi, $se, $th) = $val =~ /(\d+)-(\d+)(.*)/;
				# ' (A)' -> '(A)'
				($th) = $th =~ /\s(.*)/ if length $th > 0;
				return ($fi-1, ($se-$fi+1), $th);
			}
			default {
				($fi, $se) = $val =~ /(\d+)-(\d+)/;
				return ($fi-1, ($se-$fi+1));
			}
		}
	}
	
	sub sps_rel {
		my $old_format = shift;
		unless (defined $old_format) {
			print "";
		}
		my ($start,$last, $suffix) = $old_format =~ /(\d+)-(\d+)(.*)/;
		$suffix = '' unless defined $suffix;
		# return laenge und suffix
		return ($last-$start,$suffix);
	}
	
	# SPSS-Variablen aus .dat lesen
	#
	sub readSPSdat{
		my @datlines = @_;
		# DH1_02 188-189
		my @datVars;
		# (var,offset,length)
		# (DH1_02,187,1)
		foreach my $datline (@datlines) {
			next unless $datline =~ /^\s/;
			next unless $datline =~ /^\s\/1/; # skip bis zu den ersten Varibalen
			my $varname, my $offset = $datline =~ /^\s(.+)\s(\d.+)-(\d.+)/;
			#my $offset;
			my $length;
			push @datVars, [$varname,$offset,$length];
		}
	}
	
	# add zeros until string is a (input)-digit string
	#
	sub addzeros {
		my $pagenumber = shift;
		my $final_digits = shift;
		my $digits = length $pagenumber;
		for (my $j = $final_digits; $j > $digits; $j--) {
			$pagenumber = join('', '0', $pagenumber);
		}
		return $pagenumber;
	}
	
	sub addchar {
		my $base = shift;
		$base = '' unless defined $base;
		my $final_length = shift;
		my $char = shift;
		$char = ' ' unless defined $char;
		my $num_char = length $base;
		for (my $j = $final_length; $j > $num_char; $j--) {
			$base =  join '', $char, $base;
		}
		return $base;
	}
	
	sub path_perl {
		my $path = shift;
		$path =~ s/\\/\\\\/g;
		return $path;		
	}
	
	sub path_win {
		my $path = shift;
		$path =~ s/\\\\/\\/g;
		return $path;		
	}
	
	sub path_winTOlinux {
		# win to linux pfad-format
		my $path = shift;
		$path =~ s/\\/\//g;
		return $path;
	}
	
	sub path_linuxTOwin {
		# win to linux pfad-format
		my $path = shift;
		$path =~ s/\//\\/g;
		return $path;
	}
	
	sub specialJoin {
		my $fieldTerminator = shift;
		my $fieldEnclosed = shift;
		my @values = @_;
		#(1,3,4,asdas,12)
		#'1';'3';'4';'asdas';'12'
		return $fieldEnclosed . (join $fieldEnclosed.$fieldTerminator.$fieldEnclosed, @values) . $fieldEnclosed;
	}
	
	sub prep_bildd {
		my $prefix = shift;
		my @bildd = @_;
		my @nbildd;
		foreach my $bildd (@bildd) {
			my @n2bildd;
			foreach my $e (split /;/, $bildd) {
				
				if ($e =~ /TIF$/) {
					$e = Data::remove_ws $e;
					$e =~ s/\.TIF/\.png/;
					$e =~ tr/A-Z/a-z/;
					#$e =~ tr/\\/\//;
					$e = $prefix.$e;
				}
				push @n2bildd, $e;
			}
			push @nbildd, (join ';', @n2bildd);
			@n2bildd = ();
		}
		return @nbildd;
	}
	
	
	sub forms_count {
		my $in = shift;
		warn "ungleich acht!" if length $in != 8;
		#M;D;H;mm;iii
		#1-9ABC;1-9ABCDEFGHIJKLMNOPQRSTUV;1-9ABCDEFGHIJKLMNO(?);00-59;000-999
		my @M = (1 .. 9, 'A','B','C');
		my @D = (1 .. 9, 'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V');
		my @H = (1 .. 9, 'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O');
		my @mm = (0 .. 59);
		my @iii = (0 .. 999);
		
		my $m_in = substr $in,0,1;
		my $d_in = substr $in,1,1;
		my $h_in = substr $in,2,1;
		my $mm_in = substr $in,3,2;
		my $iii_in = substr $in,5,3;
		
		if ($iii_in eq '999') {
			$iii_in = '000';
			if ($mm_in eq '59') {
				$mm_in = '00';
				if ($h_in eq 'O') {
					$h_in = '1';
					if ($d_in eq 'V') {
						$d_in = '1';
						if ($m_in eq 'C') {
							$m_in = '1';
						} else {
							my $i = 0;
							foreach (@M) {
								if ($m_in eq $M[$i]) { $m_in = $M[$i+1]; last; }
								$i++;
							}
						}
					} else {
						my $i = 0;
						foreach (@D) {
							if ($d_in eq $D[$i]) { $d_in = $D[$i+1]; last; }
							$i++;
						}
					}
				} else {
					my $i = 0;
					foreach (@H) {
						if ($h_in eq $H[$i]) { $h_in = $H[$i+1]; last; }
						$i++;
					}
				}
			} else { $mm_in = Data::addzeros(++$mm_in, 2); }
		} else { $iii_in = Data::addzeros(++$iii_in, 3); }
		
		my $count = $m_in.$d_in.$h_in.$mm_in.$iii_in;
		
		return $count;
	}

}
1;

__END__

#########
## alt ##
#########

##################################
# log handler: warnings and dies
my $log = Log::Handler->new();
my $DOWARN = 1;
my %config = File::readfile('config_V6.csv', 'config');
my $reportdir = '';
my $reportfile = $config{error_handle_log};
BEGIN {
	use Log::Handler; 
	$SIG{'__WARN__'} = sub {
		if ($DOWARN) {
			warn @_;
			$log->warning(@_);
		}
	};
	$SIG{'__DIE__'} = sub {
		if ($DOWARN) {
			die @_;
			$log->critical(@_);
		}
	}
}
$log->add(
	file => {
		filename	=> $reportfile,
		maxlevel	=> 7,
		minlevel	=> 0,
		debug_trace	=> 1,
		debug_mode	=> 1,
		debug_skip	=> 2,
	},
);
################################


=pod

=head1 NAME

Essent.pm - wesentliche, oft gebrauchte Routinen

=head1 PACKAGE File

=head2 readfile

in: einzulesende Datei mit vollem Pfad
out: array mit Zeilen des Dateiinhalts

=head2 writefile
in: Dateiname, Schreibmodus, array
out: Datei

=head1 PACKAGE Process

=head2 confirm
in: -
out: -

Wartet auf Bestaetigung mit 'j'

=head1 PACKAGE Data

=head2 chomping
in: $variable
out: $variable ohne Zeilenumbruch


=cut
