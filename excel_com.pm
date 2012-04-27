use File::Copy;
use strict;
use warnings;
use feature qw(say switch);
use Excel_lib;
use Win32::OLE qw(in with);
use Essent;

# TODO welche Funktionen brauche ich?
use Win32::OLE::Const 'Microsoft Excel';
# TODO Welche Konstanten sind mir dadurch zugänglich?
$Win32::OLE::Warn = 3;
# The value of $Win32::OLE::Warn determines what happens when an OLE error occurs.
# If it's 0, the error is ignored. If it's 2, or if it's 1 and the script is
# running under -w, the Win32::OLE module invokes Carp::carp(). If $Win32::OLE::Warn
# is set to 3, Carp::croak() is invoked and the program dies immediately.

my $Range;

print "Modul excel_com.pm importiert.\n";

{
	package Excelobject;
	$Excelobject::VERSION = "0.5";
	#@Excelobject::ISA = qw(Range);
	
	sub new {
		my $class = shift;
		my $self = {};
		#$self -> {EXCELFILE} = $_[0];
		bless($self, $class);
		$Range = Range->new();
		
		## default-settings
		$self->{add_cell} = 1;	# add cell before data dumping
		$self->{transpose_level} = 0;	# insert formula instead of copy values
		$self->{confirm_execute} = 1;
		$self->{execute_show_all} = 0;
        $self->{check_exist} = 1;   # batch_col
		
		$Range->{add_cell} = 1;	# add cell before data dumping
		$Range->{transpose_level} = 0;	# insert formula instead of copy values
		##
		
		return $self;
	}
	
	# save copy nach Auswahl des workbooks
	sub init {
		my $self = shift;
		# use existing instance if Excel is already running
		my $excel;
		# TODO: test, wenn keine Excel-Instanz laeuft etc.
		# TODO: bestimmtes Excel-File öffnen
		# TODO: alle Excel-Threads erfassen und aufzählen/ wählen, CREATOR?
		# You can also directly attach your program to an already running OLE server:
		print "Count OLE-Objects:", Win32::OLE->EnumAllObjects(),"\n";
		eval {$excel = Win32::OLE->GetActiveObject('Excel.Application')};
        die "Excel not installed" if $@;
        # create new excel-server, returns a server object
		unless (defined $excel) {
			#$ex = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
			$excel = Win32::OLE->new('Excel.Application', 'Quit')
                    or die "Cannot start Excel";
        }
		# Excel-Server
		$self->{EXCEL} = $excel;
		
		## suche window aus?
		#my $window_count = $excel->Windows->Count;
		#if ($window_count == 0) {
		#	#$excel->Windows->Add;
		#	$excel->Windows(1)->Open;
		#	#$excel->{Visible} = 1;
		#	
		#} else {
		#	print "$window_count Windows: ";
		#	foreach (1..$window_count) {
		#		print "$_:".$excel->Windows($_)->Parent->Name.",";
		#		# oder: $ex->ActiveWindow->Caption;
		#	}
		#	print "\n";
		#	my $parent = $excel->Windows(1)->Parent->Name;
		#	my $active2 = $excel->ActiveWorkbook;
		#	print "";
		#}
		## suche workbook aus
		
		my $workb_count = $excel->Workbooks->Count;
		if ($workb_count == 0) {
			$excel->Workbooks->Add;
		} elsif ($workb_count == 1) {
			$excel->Workbooks(1)->Activate;
			print "select workbook: '", $excel->Workbooks(1)->Name, "'\n";
			$self->{WORKBOOK} = $excel->Workbooks(1);
		} else {
			print "$workb_count Workbooks: ";
			foreach (1..$workb_count) {
				print "$_:".$excel->Workbooks($_)->Name.",";
			}
			print "\n";
			my $workb_select = Process::confirm_numcount($workb_count);
			$excel->Workbooks($workb_select)->Activate;
			print "select workbook: '", $excel->Workbooks($workb_select)->Name, "'\n";
			$self->{WORKBOOK} = $excel->Workbooks($workb_select);
		}
		
		##
		## suche worksheet aus
		my $works_count = $excel->Worksheets->Count;
		print "$works_count Sheets: ";
		if ($works_count == 1) {
			$self->{WORKSHEET} = $excel->Worksheets(1);
			print "select worksheet: '", $excel->Worksheets(1)->Name, "'\n";
		} else {
			foreach (1 .. $works_count) {
				print "$_:".$excel->Worksheets($_)->Name.",";
			}
			print "\n";
			my $works_select = Process::confirm_numcount($works_count);
			## TODO: ausbauen um mit mehreren Sheets zu arbeiten
			$self->{WORKSHEET} = $excel->Worksheets($works_select);
			print "select worksheet: '", $excel->Worksheets($works_select)->Name, "'\n";
		}
		return $excel;
	}

	### getter & setter
	
	## set || get join-separator
	
	sub lastlast_row {
		my $self = shift;
		my $col = shift;
		# my $range = $self->{WORKSHEET}->Cells($caseln, $col);
		# Range("A65536").End(xlup).Select
		#my $lastlast_row = $self->{WORKSHEET}->Range($col,65536)->End('xlup')->Select;
		#$self->{WORKSHEET}->Cells($col,65536)->End('xlup')->Select;
		#my $cell = $self->{WORKSHEET}->Cells($col,65536);
		#my $LastRow = $self->{WORKSHEET}->Range("A1")->SpecialCells('xlCellTypeLastCell')->Row;
		#my $c = $self->{WORKSHEET}->Rows->Count("A");
		#my $LastRow = $self->{WORKSHEET}->Cells("A100")->SpecialCells('xlCellTypeLastCell')->Row;
		#my $lastrow = $self->{WORKSHEET}->Cells->SpecialCells('xlCellTypeLastCell')->Activate;
		my $tee = $self->{EXCEL}->XlCellType->xlCellTypeLastCell;
		
		#my $tee = $self->{WORKSHEET}->Range("A200")->Select;
		# ->End("xlUp");
		my $lastlast_row = $self->{WORKSHEET}->Range("A200")->End->{'xlUp'}->Select;
		#1048576
		return $lastlast_row;
	}
	
	sub join_sep {
		my $self = shift;
		$self->{join_sep} = shift || return $self->{join_sep};
	}
	
	## in: value-> set value
	## undef: get value
	sub activecell_val {
		my $self = shift;
		my $val = shift;
		if (defined $val) {
			$self->{EXCEL}->ActiveCell->{'Value'} = $val;
		} else {
			$self->{activecell}{value} = $self->{EXCEL}->ActiveCell->{'Value'};
		}
	}
	
	# set || get activell-position
	# TODO auch mit cells-object nutzbar
	sub activecell_pos {
		my $self = shift;
		my ($row, $col) = @_;
		if ( defined $row && defined $col) {
			$self->{EXCEL}->Cells($row, $col)->Select;
			@{$self->{activecell}{pos}} = ($self->{EXCEL}->ActiveCell->Row, $self->{EXCEL}->ActiveCell->Column);
		} else {
			@{$self->{activecell}{pos}} = ($self->{EXCEL}->ActiveCell->Row, $self->{EXCEL}->ActiveCell->Column);
		}
	}
	
	## bei Daten in Zellen einfügen zuvor Zellen neu einfügen
	## in: 0 || 1 (default)
	sub add_cell {
		my $self = shift;
		my $val = shift;
		if (defined $val) {
			$self->{add_cell} = $val;
			$Range->{add_cell} = $val;
		} else {
			return $self->{add_cell};
		}
		
	}
	
	## transpose_level 0||undef: Formelbezug
	## 1: Wert kopieren	
	sub transpose_level {
		my $self = shift;
		$self->{transpose_level} = shift || return $self->{transpose_level};
	}
	
	## set position(row, col)
	sub pos {
		my $self = shift;
		$self->{row} = shift;
		$self->{col} = shift;
	}
	
	## getter und setter für regular expression
	sub regex {
		my $self = shift;
		$self->{regex} = shift || return $self->{regex};
		my @regex = $self->{regex} =~ /(^\/)(.*)(\/$)/;
		if ( scalar @regex == 3) {
			$self->{regex} = $2;
		} else {
			warn "regular expression muss mit Slash ('/') beginnen und enden!\n";
			warn "->$self->{regex}<-\n";
		}
	}
	
	## active_cell undef: keine Relevanz
	## 'aim': Aktionsziel
	sub activecell {
		my $self = shift;
		my $keyword = shift;
		my $row = shift;
		my $col = shift;
		given ($keyword) {
			when ('aim') { @{$self->{activecell}{aim}} = ($row, $col); }
			when ('lastlastrow') {}
			default {}
		}
	}
	
	sub batch_col_block {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		if (!defined $row && !defined $col) {
			$self->activecell_pos();
			($row, $col) = @{$self->{activecell}{pos}};
		}
        my @collect_execute;
        $self->{collect_execute} = 1;
		while (defined $self->{WORKSHEET}->Cells($row, $col)->{'Value'}) {
			# $path_execute, $filename, $batch_string
            push @collect_execute, [$self->batch_col($row, $col)];
			$col++;
		}
        die "ActiveCell is empty!" unless (@collect_execute);
        $self->{collect_execute} = 0;
        
        # executes sammeln
        my $Command = Command->new;
        foreach (@collect_execute) {
            my ($path_execute, $filename, $batch_string) = @$_;
            File::writefile($filename, $batch_string);
            $Command->execute_batch($path_execute, $filename);
        }
	}
	
	## batch_row
	## lies Zellen aus und schreibe in batch
	## erste Zelle: path to execute
	## default: confirm_execute
	sub batch_col {
		my $self = shift;
		# TODO abstrahieren, bzw. intro einer Methode immer das selbe??
		my $row = shift;
		my $col = shift;
		my $filename;
		if (!defined $row && !defined $col) {
			$self->activecell_pos();
			($row, $col) = @{$self->{activecell}{pos}};
		}
		my $path_execute = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
        die "ActiveCell is empty!" unless ($path_execute);
		die "No directory: ->$path_execute<-: $!" unless (-e $path_execute);
		$row++;
		$path_execute =~ s/\\/\\\\/g;
		unless ($path_execute =~ /\\\\$/) {
			$path_execute .= "\\\\";
		}
		$filename = $path_execute."excel_batch.bat";
		my @array = $self->readcol($row, $col);
        if ($self->{check_exist}) {
            my $Guess = Guess->new();
            $Guess->{path} = $path_execute;
            @array = $Guess->parse(@array);
        }
		my $batch_string = join "\n", @array;
        if ($self->{collect_execute}) {
            return ($path_execute, $filename, $batch_string);
        } else {
            File::writefile($filename, $batch_string);
            my $Command = Command->new;
            $Command->execute_batch($path_execute, $filename);    
        }
	}
	
	## Zeile 1 in Spalten
	## Zeile 2 in Spalten darunter
	## was ist mit leeren Zellen?
	## default: read at activecell
	## write bei lastlast_row
	
	sub Zeilen_in_1Spalte {
		my $self = shift;
		my $readrow = shift;
		my $readcol = shift;
		my $writerow = shift;
		my $writecol = shift;
		$self->{transpose_level} = 0 unless defined $self->{transpose_level};
		unless (defined $readrow && defined $readcol) {
			$self->activecell_pos();
			($readrow, $readcol) = @{$self->{activecell}{pos}};
		}
		unless (defined $writerow && defined $writecol) {
			# TODO lastlast_row
			# TODO suche erste freie Zeile zum Beschreiben
			if (defined $self->{activecell}{aim}) {
				($writerow, $writecol) = @{$self->{activecell}{aim}};
			} else {	# default
				$writerow = $self->last_row($readrow,$readcol) + 1;
				$writecol = $readcol;
			}
		}
		my @vals_arraySchachtel;
		while (my $vals_ref = [ $self->readrow($readrow, $readcol) ] ) {
			say "row:".$readrow.",";
			my @vals = @{$vals_ref};
			last if ((scalar @vals) == 0);
			if ($self->{transpose_level} == 1) {
				# Values einsetzen/ kopieren
				foreach (@vals) {
					push @vals_arraySchachtel, [$self->val_format($_)];
				}
				$readrow++;
			} else {
				# Formelbezug setzen
				# Formeln ableiten
				for (my $i = 0; $i < scalar @vals; $i++) {# count rows of every row, take max for write_range
					my $sourcecell = $self->{WORKSHEET}->Cells($readrow, $readcol+$i);
					my @sourcecell_A1 = $self->R1toA1($sourcecell);
					my $cell_A1 = "=$sourcecell_A1[0]$sourcecell_A1[1]";
					push @vals_arraySchachtel, [$cell_A1];
				}
				$readrow++;
			}
		}
		my $Range = Range->new();
		$Range->{RANGE_START} = $self->{WORKSHEET}->Cells($writerow, $writecol);
		$Range->{WORKSHEET} = $self->{WORKSHEET};
		$Range->write_range(@vals_arraySchachtel);	# [[1],[2],[3],[4]];
	}
	
	sub Spalten_in_1Zeile {
		my $self = shift;
		my $readrow = shift;
		my $readcol = shift;
		my $writerow = shift;
		my $writecol = shift;
		$self->{transpose_level} = 0 unless defined $self->{transpose_level};
		unless (defined $readrow && defined $readcol) {
			$self->activecell_pos();
			($readrow, $readcol) = @{$self->{activecell}{pos}};
		}
		unless (defined $writerow && defined $writecol) {
			# TODO lastlast_row
			# TODO suche erste freie Zeile zum Beschreiben
			if (defined $self->{activecell}{aim}) {
				($writerow, $writecol) = @{$self->{activecell}{aim}};
			} else {	# default
				$writerow = $readrow;
				$writecol = $self->last_col($readcol,$readrow) + 1;
			}
		}
		my @vals_arraySchachtel;
		while (my $vals_ref = [ $self->readcol($readrow, $readcol) ] ) {
			say "col:".$readcol.",";
			my @vals = @{$vals_ref};
			last if ((scalar @vals) == 0);
			if ($self->{transpose_level} == 1) {
				# Values einsetzen/ kopieren
				foreach (@vals) {
					push @vals_arraySchachtel, [$self->val_format($_)];
				}
				$readcol++;
			} else {
				# Formelbezug setzen
				# Formeln ableiten
				for (my $i = 0; $i < scalar @vals; $i++) {	# count cols of every row, take max for write_range
					my $sourcecell = $self->{WORKSHEET}->Cells($readrow+$i, $readcol);
					my @sourcecell_A1 = $self->R1toA1($sourcecell);
					my $cell_A1 = "=$sourcecell_A1[0]$sourcecell_A1[1]";
					push @vals_arraySchachtel, [$cell_A1];
				}
				$readcol++;
			}
		}
		my $Range = Range->new();
		$Range->{RANGE_START} = $self->{WORKSHEET}->Cells($writerow, $writecol);
		$Range->{WORKSHEET} = $self->{WORKSHEET};
		$Range->write_range(@vals_arraySchachtel);	# [[1],[2],[3],[4]];
		
	}
	


	# stop bei letzter Zelle mit Content
	
	sub last_row {
		my $self = shift;
		my $startrow = shift;
		my $col = shift || 1;
		#my $row = 1;
		#my $row = 0;
		my $lastrow = $startrow;
		while (defined($self->{WORKSHEET}->Cells($lastrow+1, $col)->{'Value'} )) {
			$lastrow++;
		}
		return $lastrow;
	}
	
	sub last_col {
		my $self = shift;
		my $row = shift;
		my $startcol = shift;
		my $lastcol = $startcol;
		while (defined($self->{WORKSHEET}->Cells($row, $lastcol+1)->{'Value'} )) {
			$lastcol++;
		}
		return $lastcol;
	}
	
	# TODO auch mit Range-Object verwendbar machen?
	sub readcol {
		my $self = shift;
		my @colarray;
		my $row = shift;
		my $col = shift || 1;
		# row, column
		# TODO finde letzte Zelle und lies als Range - schneller?
		while ( defined($self->{WORKSHEET}->Cells($row, $col)->{'Value'} )) {
			push(@colarray, $self->{WORKSHEET}->Cells($row, $col)->{'Value'});
			$row++;
		}
		$self->{readcol_row} = $row;	# aktuelle Leseposition
		return @colarray;
	}

	## in: array
	## out: array of array of regex result ('04' -> ('0','4'))
	sub regex_array {
		my $self = shift;
		my @values = @_;
		my @regex_result;
		foreach my $value (@values) {
			my @regex_value = $value =~ /$self->{regex}/;
			push @regex_result, [@regex_value];
		}
		return @regex_result;	# [[0,2],[0,4],[]]
	}
	
	## regex_col
	## default:
	##  activecell = regex
	##  $row+1-> readcol
	##  Column->Add (in write_range)
	##  $col+1->write_range
	## in:
	sub regex_col {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $readrow;
		given ($self->regex) {
			when ('activecell') { $self->regex($self->activecell_val()) }
			default {$self->regex($self->activecell_val())}
		}
		warn "Kein Regular Expression definiert!\n" unless defined $self->{regex};
		if (!defined $row && !defined $col) {
			$self->activecell_pos();
			($row, $col) = @{$self->{activecell}{pos}};
			$readrow = $row+1;
		}
		#my $range_attr = $Range->range_attr();
		#if (scalar (keys %$range_attr) == 0) {	# keine settings -> default settings
		#	$col++;
		#	my $cells_start = $self->{WORKSHEET}->Cells($row, $col);
		#	$Range->range_attr($cells_start);
		#	print "";
		#}
		
		
		my @array = $self->readcol($readrow, $col);
		my @regex_result = $self->regex_array(@array);
		# add col
		# in write_range
		#$self->{WORKSHEET}->Columns($col)->Insert;
		$Range->{WORKSHEET} = $self->{WORKSHEET};
		$Range->{RANGE_START} = $Range->{WORKSHEET}->Cells($readrow, $col+1);
		$Range->write_range(@regex_result);
		return $Range->{RANGE};
	}
	
	sub write_range {
		my $self = shift;
		my @arr = @_;
		$Range->write_range(@_);
	}

	sub readrow {
		my $self = shift;
		my @rowarray;
		my $row = shift;
		my $col = shift || 1;
		my @last;
		# row, column
		# TODO row als array einmal einlesen - schneller??
		while ( defined($self->{WORKSHEET}->Cells($row, $col)->{'Value'} )) {
			push(@rowarray, $self->{WORKSHEET}->Cells($row, $col)->{'Value'});
			@last = ($row, $col);
			$col++;
		}
		$self->{readrow_col} = $col;	# letzte, nicht verwendete Leseposition
		return @rowarray;
	}
	
	sub join_row {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $sep = shift;
		unless (defined $sep) {
			# separator from $self
			$sep = $self->{join_sep};
			# wenn auch undef->default
			$sep = ',' unless defined $sep;
		}
		unless (defined $row && defined $col) {
			($row, $col) = $self->activecell_pos();
		}
		
		my @rowarray = $self->readrow($row, $col);
		my $joinedrow;
		foreach (@rowarray) {
			next if $_ eq '';
			if (defined $joinedrow) {
				$joinedrow = $joinedrow . $sep . $_;
			} else {
				$joinedrow = $_;
			}
		}
		return $joinedrow;
	}

	# TODO
	sub join_col {
		my $self = shift;
	}
	
	sub join_row_block {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		unless (defined $row && defined $col) {
			$self->activecell_pos;
			($row,$col) = @{$self->{activecell}{pos}};
		}
		my $last_row = shift || $self->last_row($row, $col);
		my $sep = $self->{join_sep} || ",";
		my $writerow = $row;
		my $writecol = 1;
		my @joined_row_array;
		while ($row <= $last_row) {
			$self->join_row($row, $col);
			# ermittle die aeußerste Spalte zum Beschreiben
			$writecol = $self->{readrow_col} if $writecol < $self->{readrow_col};
			push @joined_row_array, [$self->join_row($row, $col)];
			#$self->{WORKSHEET}->Cells($row, $self->{readrow_col}+1)->{'Value'} = $self->join_row($row, $col)
			$row++;
		}
		# TODO: write als Range-Objekt
		# default: add_col, write in  col+1
		$self->{WORKSHEET}->Columns($writecol)->Insert;
		$Range->{WORKSHEET} = $self->{WORKSHEET};
		$Range->{RANGE_START} = $Range->{WORKSHEET}->Cells($writerow, $writecol);
		$self->write_range(@joined_row_array);	
	}
	
	# TODO join_col_block
	sub join_col_block {
		my $self = shift;
	}
	
	sub val_format {
		my $self = shift;
		my $val = shift;
		$val = "=TEXT($_;\"00000\")" if ($val=~ /^0/);
		return $val;
	}
	
	# in: Excel->Cells-Objekt
	# out: Excel->Rangel-Objekt?
	# oder string
	sub R1toA1 {
		my $self = shift;
		# TODO Unterscheide zwischen input als Cells-Objekt oder Zellen-Tupel (row, col)
		my $cells_object = shift;
		my $row = $cells_object->Row;
		my $col = $cells_object->Column;
		my $col_str = Excellib::rangetocell_format($col);
		return ($col_str,$row);
		
		
		# TODO return Range-Objekt
		#my $range_object = $self->{WORKSHEET}->Range("$input0$row");
		#return $range_object;
	}
	
	## in: Attribute für regex_col
	## ...
	## out#
	sub regex_col_attr {
		my $self = shift;
		my @arr = @_;
		my $key;
		my $val;
		if (scalar @arr > 1) {	# set attr
			while (@arr) {
				$key = shift @arr;
				$val = shift @arr;
				$self->{regex_col_attr}{$key} = $val;
			}
		} elsif (scalar @arr == 1) {	# get attr of key
			$key = shift @arr;
			return $self->{regex_col_attr}{$key};
		} else {	# get all attr as hash
			return %{$self->{regex_col_attr}};
		}
	}
#########
## alt ##
#########

	# mit Handling für Zahlen
	sub writeval {
		my $self = shift;
		my $val = shift;
		my $insert;
		if ($val=~ /^0/) {
			$insert = "=TEXT($val;\"00000\")";
		} else {
			$insert = $val;
		}
		$self->{WORKSHEET}->Cells($self->{row},$self->{col})->{'Value'} = $insert;
	}

	sub bildd_dirs {
		my $self = shift;
		if (@_) {
			$self->{bildd_dir_rel} = $_[0] if ($_[0]);
			$self->{bildd_dir_nest} = $_[1] if ($_[1]);
		}
		return ($self->{bildd_dir_rel}, $self->{bildd_dir_nest});
	}
	
	sub add_head {
		my $self = shift;
		my @head = @_;
		$self->{var_sort} = \@head;
		# für insert_case
		for (my $i = 0; $i < scalar @head; $i++) {
			print "";
			$self->{WORKSHEET}->Cells(1,$i+1)->{'Value'} = $head[$i];
		}
	}
	
	# lege Variablenformat an (SPSS -> substr-Format)
	sub var_format_tosubstr {
		my $self = shift;
		#print "ref_varfromat", ref $_[0], "\n";
		my %vars = %{$_[0]};
		my %vars_new;
		foreach my $var (keys %vars) {
			# key		val
			# FrB01		'9-13'
			#$vars_new{$var} = Data::spss_position($vars{$var});
			(my $start, my $le) = Data::spss_position($vars{$var});
			$vars_new{$var}{'start'} = $start;
			$vars_new{$var}{'le'} = $le;
			# {'FrB01'} = (9,5);
		}
		$self->{var_format} = {%vars_new};
	}
	
	sub adjust_cells {
		my $self = shift;
		# Width < 400, sonst Ausnahmefehler
		$self->{WORKSHEET}->Columns(9)->{'ColumnWidth'} = 65;
		#$self->{WORKSHEET}->Rows(1)->{'RowHeight'} = 50;
	}
	
	# füge Werte von case in excel-sheet ein
	sub insert_case {
		my $self = shift;
		#$self->{WORKSHEET}->Columns(9)->{'ColumnWidth'} = 250;
		#$self->{WORKSHEET}->Rows(1)->{'RowHeight'} = 50;
		my $caseln = shift;
		my $case = shift;
		# arbeite hash ab
		my $col = 1;
		# Reihenfolge!
		my $bildd1;
		my $schnipp;
		foreach my $var (@{$self->{var_sort}}) {
			my $val_start = ${$self->{var_format}}{$var}{'start'};
			my $val_le = ${$self->{var_format}}{$var}{'le'};
			my $val = substr $case, $val_start, $val_le;
			
			#$bildd1 = $val if ($var eq 'Bildd');
			given ($var) {
				when (/Importfi/) {
					$val = Data::extractfilename($val);
					#$bildd1 = $val;
					# Chwa: Erweiterung von tif zu png umbauen
					#$bildd1 = $val =~ /(.+\.)tif/ . "png";
					$val  =~ s/tif$/png/;
					$bildd1 = $val;
				}
				when (/Fieldfil/) {
					$val = Data::extractfilename($val);
					$schnipp = $val;
				}
				# Chwa-option
				# SS01.p00000003.tif
				when (/Bild_Seite3/) {
					my $pref = substr $val, 0, 14;
					$schnipp = join('', $pref, '_kennw.png');
				}
			}
			
			$self->{WORKSHEET}->Cells($caseln,$col)->{'Value'} = $val;
			$col++;
		}
		
		# insert softlink (relative path)
#		my $range = $self->{WORKSHEET}->Cells($caseln, $col);
#		print "bildd:", $bildd1, "-\n";
#		my $bildd1_rel = join('', $self->{bildd_dir_rel}, $bildd1);
#		$self->{WORKSHEET}->Hyperlinks->Add({
#			Anchor => $range,
#			Address => $bildd1_rel,
#			TextToDisplay => $bildd1_rel,
#			ScreenTip => "Bilddatei oeffnen mit Mausklick"
#		});
		
		# insert $bildd1 as nested (thumbnail)
		my $schnipp_abs = join('', $self->{bildd_dir_nest}, $bildd1);
		$self->{WORKSHEET}->Rows($caseln)->{'RowHeight'} = 55;
		my $top_align = ($caseln-2)*54.75+20;
		$self->{WORKSHEET}->Shapes->AddPicture (
			$schnipp_abs,		# Filename As String
			1,					# LinkToFile As MsoTriState
			1,					# SaveWithDocument As MsoTriState
			490,				# Left As Single
			$top_align,	# Top As Single
			350,				# Width As Single
			40					# Height As Single
		);
		print "";
		
#		my $lastrow = $self->{WORKSHEET}->
#		         nLastRow = .Cells.Find(What:="*", After:=.Cells(1), _
#                LookIn:=xlFormulas, LookAt:=xlWhole, _
#                SearchOrder:=xlByRows, _
#                SearchDirection:=xlPrevious).Row
		
	}
	
	sub init_Excelfile {
		my $self = shift;
		my $file = shift;
		my $sheet = shift;
		
		##
		my $Count = Win32::OLE->EnumAllObjects(sub {
			my $Object = shift;
			my $Class = Win32::OLE->QueryObjectType($Object);
			printf "# Object=%s Class=%s\n", $Object, $Class;
		});
		print $Count;
		##
		my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
			|| Win32::OLE->new('Excel.Application', 'Quit');
		#my $Book = $Excel->Workbooks->Add($_[0]);
		my $Book = $Excel->Workbooks->Open($file);
		#$self->{WORKSHEET} = $Book->Worksheets($self->{WORKSHEETNR});
		$self->{WORKSHEET} = $Book->Worksheets($sheet);	
	}
}

{
    package Guess;
    $Guess::VERSION = "0.1";
    #@Guess::ISA = qw(Excelobject);
    
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
        #$self->{lib} = qw(copy move del);
        %{$self->{lib}} = (copy=>"group2",
                           move=>"group2",
                           del=>"group1");
		return $self;
	}
    
    sub parse {
        my $self = shift;
        my @array = @_;
        my @parsed_array;
        # TODO alternative mit for oder foreach...
        for (my $i = 0; $i < scalar @array; $i++) {
            my @kes = keys %{$self->{lib}};
            if ( my ($lib) = grep {$array[$i] =~ /^($_)/i} @kes ) {
                # $1 nicht ?!
                my $hit = ${$self->{lib}}{$lib};
                $array[$i] = $self->$hit($array[$i]);
                #print "";
            }
        }
        return @array;
    }
    
    ## group1: DEL
    sub group1 {
        my $self = shift;
        my $string = shift;
        $string =~ /()/;
    }
    
    ## group2: COPY MOVE
    sub group2 {
        my $self = shift;
        my $string = shift;
        # COPY 2012-03-15 Vereinsliste aktive.xlsx HIGH2 tee.txt
        my $path;
        # TODO: grab out of $self->{hit}
        # and put in regex-match!
        my @hits = qw(copy move);
        my $regex_hits = join "|", @hits;
        # remove hit
        #$string =~ /($regex_hits)\s(\w+.*)/i;
        $string =~ /(copy|move)\s(\w+.*)/i;
        my $hit = $1;
        $string = $2;
        my ($file1, $file2) = $self->rebuild($string);
        return "$hit $file1 $file2";
        # TODO
        # COPY 2012-03-15 Vereinsliste aktive.xlsx tee.txt
        # Endergebnis:
        # (COPY) (2012-03-15 Vereinsliste aktive.xlsx) (tee.txt)
        # mit " umrahmen: COPY "2012-03-15 Vereinsliste aktive.xlsx" tee.txt
        # return
    }
    
    sub rebuild {
        # '2012-03-15'
        # '2012-03-15 Vereinsliste'
        # '2012-03-15 Vereinsliste aktive.xlsx' -> passt
        my $self = shift;
        my $string = shift;
        my $ws = 0;
        my @string_split = split / /, $string;
        my $string_rebuild;
        my $string_rest;
        my $file_exist = 0;
        # TODO for (@string_split) { # funzt nicht einfach so? Liest jedes mal die Anzahl Elemente?!
        my $i = 0;
        for (@string_split) {
            # $string_rebuild .= shift @string_split;
            $i++;
            $string_rebuild .= $_;
            if (-e $self->{path}.$string_rebuild) {
                $file_exist = 1;
                $string_rest = join " ", @string_split[$i..$#string_split];
                last;
            }
            $string_rebuild .= " ";
            $ws = 1;
        }
        die "first argument (file) does not exist!: ->$self->{path}.$string_rebuild<-" unless $file_exist;
        die "second argument (destination) missing\n" unless $string_rest;
        $string_rest = '"'.$string_rest.'"' if $string_rest =~ /\s/;
        $string_rebuild = '"'.$string_rebuild.'"' if $ws;
        return ($string_rebuild, $string_rest);
    }
    
}

{
	#####
    ## package Command
    ## communication with (windows) system
    #####
    
    package Command;
	$Command::VERSION = "0.1";
    # base class: Excelobject (method-scope)
    @Command::ISA = qw(Excelobject);
    
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
		return $self;
	}
    
    # TODO sub in sub funktioniert nicht!?
	sub execute_batch {
		my $self = shift;
		my $path_execute = shift;
		my $filename = shift;
		#my $result_execute = '';
		#my $result_error = '';
		#chdir($path_execute) or die "Can't change directory to $path_execute: $!";
		if ($self->{confirm_execute}) {
			print "Execute ", $filename, "?\n";
			if (Process::confirmJN()) {
				$self->execute($filename);
			} else {
				print "File not executed\n";
			}
		} else {
			$self->execute($path_execute, $filename);
			print "";
		}
		
		sub execute {
			my $self = shift;
			my $path_execute = shift;
			my $filename = shift;
			my $result_execute = '';
			my $result_error = '';
			my @operation_ok = ("1 Datei.+kopiert", "1 file.+copied");
			# TODO funktioniert trotz absolutem Pfad nicht ohne chdir?
			chdir($path_execute) or die "Can't change directory to $path_execute: $!";
			$result_execute = `$filename`;
			if ($self->{execute_show_all}) {
				print "___EXECUTE LOG: $filename ___\n";
				print $result_execute;
				print "_________________\n";
				File::writefile($filename.".log", $result_execute);
			} else {	# show only errors
				my $first;
				foreach my $result_line (split /\n/, $result_execute) {
                    next unless $result_line;   # das gleiche, wie: next if length $result_line == 0;
					if ($first) {
						if (grep {$result_line =~ /$_/} @operation_ok) {
							$first = undef;
						} else {
							$result_error .= $first."\n";
							$result_error .= $result_line."\n";
							$first = undef;
						}
					} else {
						$first = $result_line;
					}						
				}
				if ($result_error) {
					print "\n______EXECUTE ERROR: $filename ___\n";
					print $result_error;
					print "_________________________\n";
					File::writefile($filename."_ERROR.log", $result_error);
				} else {
                    print "EXECUTE OK: $filename\n";
                }
			}
		}
	}
}

{
	package Range;
	# require Exporter;
	# @ISA = qw(Exporter);
	# @EXPORT = qw(guess_media_type media_suffix);
	# @EXPORT_OK = qw(add_type add_encoding read_media_types);
	$Range::VERSION = "0.1";
	#Range->VERSION(0.1);
	# nötig?
	# ja, damit auch die Excelobjekte gehandhabt werden können (Cells, Range etc.)
	# TODO Exceobject-Variablen nicht durch Vererbung in Range verfügbar?
	@Range::ISA = qw(Excelobject);
	#use base 'Excelobject'; # sets @MyCritter::ISA = ('Critter');
	
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
		#Package::Name->can('function')
		my $cani = $self->can('join_row');
		return $self;
	}
	
	sub range {
		my $self = shift;
		$self->{RANGE} = shift || return $self->{RANGE};
	}
		#	%{$self->{attr}} = (
		#	PrintError => 1,
		#	RaiseError => 0
		#);
	

	## TODO: erste und letzte Cell von Range in Farbe
	## write_range
	## braucht: $Range->{RANGE_START}
	sub write_range {
		my $self = shift;
		my @arrofarr = @_;
		my $depth_arr = 0;
		my $range_start_row = $self->{RANGE_START}->Row;
		my $range_start_col = $self->{RANGE_START}->Column;
		# wie tief ist @arrofarr?
		# TODO mit map:
		foreach my $arr (@arrofarr) {
			my @arr_de = @$arr;
			$depth_arr = scalar @arr_de if ( (scalar @arr_de) > $depth_arr );
		}
		# $depth_arr mind. = 1: wenn @arrofarr leer, trotzdem extra Zellen (col) einfügen
		$depth_arr = 1 if $depth_arr < 1;
		# add cells
		# add cols
		if ($Range->{add_cell}) {
			for (my $i = 1; $i <= $depth_arr; $i++) {
				$self->{WORKSHEET}->Columns($range_start_col)->Insert;
			}
		}
		
		# Range berechnen: RANGE_START + Dimensionen von @arrofarr
		#my $range_end_col = $self->{RANGE_START}->Column + $depth_arr - 1;
		my $range_end_row = $range_start_row + scalar @arrofarr-1;
		my $range_end_col = $range_start_col + $depth_arr - 1;
		# $range_end_col & _row darf nie kleiner sein als $range_start_row $ _row
		$range_end_col = $range_start_col if $range_end_col < $range_start_col;
		$range_end_row = $range_start_row if $range_end_row < $range_start_row;
		my $range_start = $self->{WORKSHEET}->Cells($range_start_row, $range_start_col);
		my $range_end = $self->{WORKSHEET}->Cells($range_end_row, $range_end_col);
		#my $range = $self->{WORKSHEET}->Range($self->{RANGE_START},$range_end);
		# oder public?
		$self->{RANGE} = $self->{WORKSHEET}->Range($range_start,$range_end);
		
		# $self->{regex} erreichbar?
		#if ( $self->{regex} eq 'addcell' ) {
		if ( 0 ) {	
			# default
			# add row $or column at $position
		} else {
			# Werte einsetzen
			# [[1a,1b,1c,1d],[2a,2b,2c,2d],[3a,3b,3c]];
			$self->{RANGE}->{'Value'} = [@arrofarr];
			# Farbe für erste und letzte Zelle
			$range_start->Interior->{'ColorIndex'} = 20;
			$range_end->Interior->{'ColorIndex'} = 20;
		}
	}
	
	#sub range_attr {
	#	my $self = shift;
	#	$self->{RANGE_START} = shift || return $self->{RANGE_START};
	#	#my $attr = shift;
	#	#my $val = shift;
	#	#if (defined $val) {	# set attr
	#	#	$self->{attr}{$attr} = $val
	#	#} elsif (defined $attr) {	# get val of $attr
	#	#	return $self->{attr}{$attr};
	#	#} else {	# get all attr
	#	#	#${$self->{attr}}{hey} = "ho";
	#	#	if ($self->{attr}) {
	#	#		return $self->{attr};
	#	#	} else {
	#	#		return undef;
	#	#	}
	#	#}
	#}
	
}

{
	package Complete_Excel;
	@Complete_Excel::ISA = qw(Excelobject);
	
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
		return $self;
	}
	
	sub var_format_spss {
		my $self = shift;
		$self->{vars_spss} = shift;
		$self->{Importfi} = $self->{vars_spss}{'Importfi'};
		$self->{Fieldfil} = $self->{vars_spss}{'Fieldfil'};
	}
	
	sub fetch_excelval {
		my $self = shift;
		my $row = shift;	 
		my $col = shift;
		#$col = 7 if ($col eq 'default');
		my $val = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
		return $val;
	}
	
	sub add_excelval {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $val = shift;
		$self->{WORKSHEET}->Cells($row, $col) -> {'Value'} = $val;
	}
	
	# gehe col 7 (itemnum) runter
	# gleiche ab mit keys
	# ergänze mit vals
	# füge sps-format von Importfi
	sub complete_sheet {
		my $self = shift;
		my $row_itemnum = 4;
		my $col_itemnum = 7;
		my $last_row_in_col = $self->last_row($row_itemnum, $col_itemnum);
		#print "lastrow:", $last_row_in_col, "\n";
		for ($row_itemnum = 4; $row_itemnum < $last_row_in_col; $row_itemnum++) {
			my $excel_val = $self->fetch_excelval($row_itemnum, $col_itemnum);
			if ( exists($self->{vars_spss}{$excel_val}) ) {
				my $hash_val = $self->{vars_spss}{$excel_val};
				$self->add_excelval($row_itemnum, $col_itemnum+1, $hash_val);
				# füge in jede Zeile die Position von Importfi (bildd) ein.
				$self->add_excelval($row_itemnum, $col_itemnum+2, $self->{Importfi}) if (defined($self->{Importfi}));
				# füge in jede Zeile die Position von Fieldfil (Schnipsel) ein.
				$self->add_excelval($row_itemnum, $col_itemnum+3, $self->{Fieldfil}) if (defined($self->{Fieldfil}));
			}
		}
	}
}

{
	# Read-Version-Structure Objekt
	package RVS;
	@RVS::ISA = qw(Excelobject);
	
	sub new {
		my $class = shift;
		my $self = {};
		#$self -> {WORKSHEETNR} = 3;	# Deutsch-Version
		bless($self, $class);
		return $self;
	}
	sub set_start {
		my $self = shift;
		$self->{startrow} = shift;
		$self->{startcolumn} = shift;
	}
	sub read_version_structure {
		my $self=shift;
		my $version = $self->{WORKSHEET}->Cells(4, 2)->{'Value'};
		my $maxseiten =  $self->{WORKSHEET}->Cells(4, 1)->{'Value'};
		# aufgabenblock
		my @ab = $self->readcol(4,3);
		# defseite
		my @defseite = $self->readcol(4,4);
		# seitenum
		my @seitenum = $self->readcol(4,5);
		# itemnum
		my @itemnum = $self->readcol(4,7);
		my @itemErgebnis_pos = $self->readcol(4,8);
		my @bildd_pos = $self->readcol(4,9);
		my @itemSchnipsel_pos = $self->readcol(4,10);
		return (\$version, \$maxseiten, \@ab, \@defseite, \@seitenum, \@itemnum, \@itemErgebnis_pos, \@bildd_pos, \@itemSchnipsel_pos);
	}
}


{
	package CSV;
	@CSV::ISA = qw(RVS);
	
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
		return $self;
	}
	
	# @ver_cols = (\$version, \$maxseiten, \@ab, \@defseite, \@seitenum, \@itemnum, \@itemErgebnis_pos, \@bildd_pos, \@itemSchnipsel_pos);
}




__END__

## ausgesondert
		## suche window aus?
		#my $window_count = $ex->Windows->Count;
		#if ($window_count == 0) {
		#	$ex->Windows(1)->Open;
		#	#$ex->Windows->Add;
		#} else {
		#	print "$window_count Windows: ";
		#	foreach (1..$window_count) {
		#		print "$_:".$ex->Windows($_)->Parent->Name.",";
		#		# oder: $ex->ActiveWindow->Caption;
		#	}
		#	print "\n";
		#	my $parent = $ex->Windows(1)->Parent->Name;
		#	my $active2 = $ex->ActiveWorkbook;
		#	print "";
		#}
		## suche workbook aus
##

###########
## learn ##
###########
$server->{Bar} = $value;
$value = $server->{Bar};
# assign value of $server
$value = valof $server;
# das selbe!?
$value = $server->Invoke('');

# Objekt-Verschachtelung abkürzen
my $workbook = $excel->Workbooks;
my $workb_count = $workbook->Count;

# with
with($Chart, HasLegend => 0, HasTitle => 1);


##
Range("A1") Zelle A1 
Range("A1:B5") Zellen A1 bis B5 
Range("C5:D9;G9:H16") Eine Mehrfachmarkierung eines Bereichs 
Range("A:A") Spalte A 
Range("1:1") Zeile 1 
Range("A:C") Spalten A bis C 
Range("1:5") Zeilen 1 bis 5 
Range("1:1;3:3;8:8") Zeilen 1, 3 und 8 
Range("A:A;C:C;F:F") Spalten A, C und F 
##

## my $form = $self->{EXCEL}->ConvertFormula("formula:=$inputFormula,fromReferenceStyle:=xlR1C1,toReferenceStyle:=xlA1");
#####

    # program continues

    =begin comment text

    all of this stuff

    here will be ignored
    by everyone

    =end comment text

    =cut

1;

