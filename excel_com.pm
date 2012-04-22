use File::Copy;
use strict;
use warnings;
use feature qw(say switch);
#use Essent;

use Win32::OLE qw(in with);

# TODO welche Funktionen brauche ich?
use Win32::OLE::Const 'Microsoft Excel';
# TODO Welche Konstanten sind mir dadurch zugänglich?
$Win32::OLE::Warn = 3;
# The value of $Win32::OLE::Warn determines what happens when an OLE error occurs.
# If it's 0, the error is ignored. If it's 2, or if it's 1 and the script is
# running under -w, the Win32::OLE module invokes Carp::carp(). If $Win32::OLE::Warn
# is set to 3, Carp::croak() is invoked and the program dies immediately.

print "Modul excel_com2.pm importiert.\n";

{
	package Excelobject;
	#@Excelobject::ISA = qw(Range);
	
	my $Range;
	
	sub new {
		my $class = shift;
		my $self = {};
		#$self -> {EXCELFILE} = $_[0];
		bless($self, $class);
		$Range = Range->new();
		return $self;
	}
	
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
			my $workb_select = main::confirm_numcount($workb_count);
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
			my $works_select = main::confirm_numcount($works_count);
			## TODO: ausbauen um mit mehreren Sheets zu arbeiten
			$self->{WORKSHEET} = $excel->Worksheets($works_select);
			print "select worksheet: '", $excel->Worksheets($works_select)->Name, "'\n";
		}
		return $excel;
	}

	## transpose_level 0||undef: Formelbezug
	## 1: Wert kopieren
	sub transpose_level {
		my $self = shift;
		$self->{transpose_level} = shift || return $self->{transpose_level};
	}
	
	## active_cell undef: keine Relevanz
	## 'aim': Aktionsziel
	sub active_cell {
		my $self = shift;
		my $keyword = shift;
		#$self->{WORKSHEET}->Activate;
		my $cells_object = $self->{WORKSHEET}->ActiveCell;
		my $row = $cells_object->Row;
		#$self->{WORKSHEET}->Cells($readrow, $readcol+$i);
		#my $row = $cells_object->Row;
		#Cells($row, $col)->{'Value'}
		my $col = $self->{WORKSHEET}->ActiveCell->Select->Col;
		if ($keyword eq 'aim') {
			@{$self->{active_cell}->{aim}} = ($row, $col);
		} elsif ($keyword eq 'lastlastrow') {
			# not functionable yet
			#@{$self->{active_cell}->{lastlastrow}} = $self->lastlast_row();
		} else {
			# default
		}
		
	}
	
	## Zeile 1 in Spalten
	## Zeile 2 in Spalten darunter
	## was ist mit leeren Zellen?
	sub Zeilen_in_1Spalte {
		my $self = shift;
		my $readrow = shift;
		my $readcol = shift;
		my $writerow = shift;
		my $writecol = shift;
		unless (defined $writerow && defined $writecol) {
			my $key = $self->{active_cell};
			$writerow = ${$self->{active_cell}->{aim}}[0];
			$writecol = ${$self->{active_cell}->{aim}}[1];
			#$writerow = $self->lastlast_row($readcol)+1;
			#$writecol = $readcol;
		}
		# TODO: suche erste freie Zeile zum Beschreiben
		my @vals_n;
		#my $vals_count = 0;
		while (my $vals_ref = [ $self->readrow($readrow, $readcol) ] ) {
			say "row:".$readrow.",";
			my @vals = @{$vals_ref};
			#my $vals_count = scalar @vals;
			#$vals_count += scalar @vals;
			last if ((scalar @vals) == 0);
			
			if ($self->{transpose_level}) {
				# Values einsetzen/ kopieren
				foreach (@vals) {
					my $insert;
					if ($_=~ /^0/) {
						$insert = "=TEXT($_;\"00000\")";
					} else {
						$insert = $_;
					}
					push @vals_n, [$insert];
				}
				$readrow++;
			} else {
				# Formelbezug setzen
				# Formeln ableiten
				
				# funzt auch?
				# for (@vals) {
				for (my $i = 0; $i < scalar @vals; $i++) {
					my $sourcecell = $self->{WORKSHEET}->Cells($readrow, $readcol+$i);
					my @sourcecell_new = $self->R1toA1($sourcecell);
					# =A1
					my $cell_A1 = "=$sourcecell_new[0]$sourcecell_new[1]";
					push @vals_n, [$cell_A1];
					
					## TODO range_new als Excel-Range-Objekt!!?
					## convert: from R1C1-style to A1-style
					#my @range_new = $self->R1toA1($cells1);
					## =A1
				}
				$readrow++;
			}
		}
		
		#my $cells_start = $self->{WORKSHEET}->Cells($writerow,$writecol);
		#my $cells_end = $self->{WORKSHEET}->Cells($writerow+(scalar @vals_n)-1,$writecol);
		#$self->{range} = $self->{WORKSHEET}->Range($cells_start,$cells_end);
		#$self->{range}->{'Value'} = [@vals_n];	# [[1],[2],[3],[4]];
		
		# TODOTODO
		# Idee: Werte nur für Methode zugreifbar
		# neues Package
		# ISA?
		# ohne blessing?
		my $Range = Range->new();
		$Range->{RANGE_START} = $self->{WORKSHEET}->Cells($writerow, $writecol);
		#$self->write_range->{RANGE_START} = $self->{WORKSHEET}->Cells($writerow, $writecol);
		# und auch möglich mit tupel:
		#$self->write_range->{RANGE_START} = ($writerow, $writecol);
		# TODO Exceobject-Variablen nicht durch Vererbung in Range verfügbar?
		$Range->{WORKSHEET} = $self->{WORKSHEET};
		$Range->write_range(@vals_n);	# [[1],[2],[3],[4]];
	}
	
	sub get_WORKSHEET {
		my $self = shift;
		return $self->{WORKSHEET};
	}
	
	sub pos {
		my $self = shift;
		$self->{row} = shift;
		$self->{col} = shift;
	}
	
	sub writecol {
		my $self = shift;
		# als Range rein packen??
	}
	
	sub Spalten_in_1Zeile {
		my $self = shift;
		
	}
	
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
		
		
		#my $cell = $self->{WORKSHEET}->Cells($col,100)->Select;
		#$self->{WORKSHEET}->Cells(200,1)->Select;
		#$self->{WORKSHEET}->Cells(200,1)->End('xlUp')->Select;
		my $tee = $self->{WORKSHEET}->Range("A200")->Select;
		# ->End("xlUp");
		my $lastlast_row = $self->{WORKSHEET}->Range("A200")->End->{'xlUp'}->Select;
		#1048576
		return $lastlast_row;
	}

	# stop bei erster leeren Zelle
	sub last_row {
		my $self = shift;
		my $startrow = shift;
		my $col = shift;
		my $i = 1;
		while (defined($self->{WORKSHEET}->Cells($startrow + $i, $col)->{'Value'} )) {
			$i++;
		}
		#return ($i+$startrow, $col);
		return $i+$startrow;
	}
	
	sub last_col {
		my $self = shift;
		my $row = shift;
		my $startcol = shift;
		my $i = 1;
		while (defined($self->{WORKSHEET}->Cells($row, $startcol+$i)->{'Value'} )) {
			$i++;
		}
		return ($row, $i+$startcol);
	}
	
	sub readcol {
		my $self = shift;
		my @colarray;
		my $col = shift;
		my $row = shift || 1;
		# row, column
		#print "readcol:";#, $self->{WORKSHEET}->Cells($row, $col)->{'Value'}, "\n";
		while ( defined($self->{WORKSHEET}->Cells($row, $col)->{'Value'} )) {
			push(@colarray, $self->{WORKSHEET}->Cells($row, $col)->{'Value'});
			$row++;
		}
		return @colarray;
	}
	
	# undef: default
	# 'addcell': neue Zeile einfügen, in Form von column oder row
	sub regex {
		my $self = shift;
		$self->{regex} = shift || return $self->{regex};

	}
	
	# getter und setter für regular expression
	sub regexp {
		my $self = shift;
		$self->{regexp} = shift || return $self->{regexp};
		# brackets ('/') entfernen
		$self->{regexp} =~ s/^\/(.*)\/$/$1/;
		print "";
	}
	
	## in: array (row or column)
	## out: array of array of regex result ('04' -> ('0','4'))
	sub regex_array {
		my $self = shift;
		my @values = @_;
		## TODO
		## my $regex_in = $self->{WORKSHEET}->Cells(~pos)->{'Value'};
		## my $regex = join('', '(.*\.', $regex_in, '$)');
		my @regex_result;
		my $regexp = $self->{regexp};
		foreach my $value (@values) {
			my @regex_value = $value =~ /$regexp/;
			push @regex_result, [@regex_value];
		}
		return @regex_result;	# [[0,2],[0,4],[]]
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
	
	## regex_col
	## default:
	##  activecell = regexp
	##  $row+1-> readcol
	##  Column->Add
	##  $col+1->write_range
	## in:
	sub regex_col {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		
		given ($self->regexp) {
			when ('activecell') { $self->regexp($self->read_activecell()) }
			default {$self->regexp($self->read_activecell())}
		}
		warn "Kein Regular Expression definiert!\n" unless defined $self->{regexp};
		if (!defined $row && !defined $col) {
			($row, $col) = @{$self->{activecell}{pos}};
			$row++;
		}
		
		my @array = $self->readcol($col, $row);
		my @regex_result = $self->regex_array(@array);
		my $range_attr = $Range->range_attr();
		if (scalar (keys %$range_attr) == 0) {	# keine settings -> default settings
			my $cells_start = $self->{WORKSHEET}->Cells($row, ++$col);
			$Range->range_attr($cells_start);
			print "";
		}
		# add col
		# ActiveCell.EntireColumn.Insert
		# Workbooks("yourworkbook").worksheets("theworksheet").Columns(x).Insert
		#$self->{WORKSHEET}->ActiveCell->EntireColumn->Insert;
		$self->{WORKBOOK}->Worksheets(1)->Columns($col)->Insert;
		$Range->{WORKSHEET} = $self->{WORKSHEET};
		$Range->{RANGE_START} = $Range->{WORKSHEET}->Cells($row, $col);
		$Range->write_range(@regex_result);
		return $Range->{RANGE};
	}
	
	sub read_activecell {
		my $self = shift;
		$self->activecell();
		return $self->{activecell}{value};
	}
	
	## in: value-> set value
	##  undef: get value und pos
	sub activecell {
		my $self = shift;
		my $val = shift;
		if (defined $val) {
			$self->{EXCEL}->ActiveCell->{'Value'} = $val;
		} else {
			my $row = $self->{EXCEL}->ActiveCell->Row;
			my $col = $self->{EXCEL}->ActiveCell->Column;
			@{$self->{activecell}{pos}} = ($row, $col);
			$self->{activecell}{value} = $self->{EXCEL}->ActiveCell->{'Value'};
		}
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
		# TODO row als array einmal einlesen !?
		while ( defined($self->{WORKSHEET}->Cells($row, $col)->{'Value'} )) {
			push(@rowarray, $self->{WORKSHEET}->Cells($row, $col)->{'Value'});
			$self->{readrow_col} = $col;
			@last = ($row, $col);
			$col++;
		}
		## letzte Leseposition
		# geht auch $self->readrow->{last} ?
		return @rowarray;
	}
	
	sub join_row {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $sep = shift || ",";
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
		$self->{WORKSHEET}->Cells($row, $self->{readrow_col}+1)->{'Value'} = $joinedrow;
		return $joinedrow;
	}
	sub set_join_sep {
		my $self = shift;
		$self->{join_sep} = shift;
		#neu: $self->{join_sep} = shift || return $self->{join_sep};
	}
	sub get_join_sep {
		my $self = shift;
		return $self->{join_sep};
	}
	
	sub join_row_block {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $last_row = shift || $self->last_row($row, $col);
		my $sep = $self->{join_sep} || ",";
		while ($row <= $last_row) {
			$self->join_row($row, $col);
			$row++;
		}
	}

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
		#$self->{WORKSHEET}->Cells($self->{row},$self->{col})->{'Value'} = "=TEXT($val;\"00000\")"; 
	}
	
	
	# in: Excel->Cells-Objekt
	# out: Excel->Rangel-Objekt?
	# oder string
	#sub rangetocell {
	sub R1toA1 {
		my $self = shift;
		# TODO Unterscheide zwischen input als Cells-Objekt oder Zellen-Tupel (row, col)
		my $cells_object = shift;
		my $row = $cells_object->Row;
		my $col = $cells_object->Column;
		my $col_str = $self->rangetocell_format($col);
		return ($col_str,$row);
		
		
		# TODO return Range-Objekt
		#my $range_object = $self->{WORKSHEET}->Range("$input0$row");
		#return $range_object;
		
	}
	
	sub rangetocell_format {
		my $self = shift;
		my $input = shift;
		my %rangetocell_lib = (
			1=>"A",
			2=>"B",
			3=>"C",
			4=>"D",
			5=>"E",
			6=>"F",
			7=>"G",
			8=>"H",
			9=>"I",
			10=>"J",
			11=>"K",
			12=>"L",
			13=>"M",
			14=>"N",
			15=>"O",
			16=>"P",
			17=>"Q",
			18=>"R",
			19=>"S",
			20=>"T",
			21=>"U",
			22=>"V",
			23=>"W",
			24=>"X",
			25=>"Y",
			26=>"Z",
			27=>"AA",
			28=>"AB",
			29=>"AC",
			30=>"AD",
			31=>"AE",
			32=>"AF",
			33=>"AG",
			34=>"AH",
			35=>"AI",
			36=>"AJ",
			37=>"AK",
			38=>"AL",
			39=>"AM",
			40=>"AN",
			41=>"AO",
			42=>"AP",
			43=>"AQ",
			44=>"AR",
			45=>"AS",
			46=>"AT",
			47=>"AU",
			48=>"AV",
			49=>"AW",
			50=>"AX",
			51=>"AY",
			52=>"AZ");
		my $output = $rangetocell_lib{$input};
		return $output;
	}
#########
## alt ##
#########
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
	package Range;
	# nötig?
	# ja, damit auch die Excelobjekte gehandhabt werden können (Cells, Range etc.)
	@Range::ISA = qw(Excelobject);
	
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
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
	
	sub range_attr {
		my $self = shift;
		$self->{RANGE_START} = shift || return $self->{RANGE_START};
		#my $attr = shift;
		#my $val = shift;
		#if (defined $val) {	# set attr
		#	$self->{attr}{$attr} = $val
		#} elsif (defined $attr) {	# get val of $attr
		#	return $self->{attr}{$attr};
		#} else {	# get all attr
		#	#${$self->{attr}}{hey} = "ho";
		#	if ($self->{attr}) {
		#		return $self->{attr};
		#	} else {
		#		return undef;
		#	}
		#}
	}

	# $self->write_range->{RANGE_START}
	sub write_range {
		my $self = shift;
		my @arrofarr = @_;
		my $depth_arr = 0;
		# wie tief ist @arrofarr?
		# TODO mit map:
		foreach my $arr (@arrofarr) {
			my @arr_de = @$arr;
			$depth_arr = scalar @arr_de if ( (scalar @arr_de) > $depth_arr );
		}
		# Range berechnen: RANGE_START + Dimensionen von @arrofarr
		my $range_end_row = $self->{RANGE_START}->Row + scalar @arrofarr-1;
		my $range_end_col = $self->{RANGE_START}->Column + $depth_arr - 1;
		my $range_end = $self->{WORKSHEET}->Cells($range_end_row, $range_end_col);
		#my $range = $self->{WORKSHEET}->Range($self->{RANGE_START},$range_end);
		# oder public?
		$self->{RANGE} = $self->{WORKSHEET}->Range($self->{RANGE_START},$range_end);
		
		# $self->{regex} erreichbar?
		#if ( $self->{regex} eq 'addcell' ) {
		if ( 0 ) {	
			# default
			# add row $or column at $position
		} else {
			# Werte einsetzen
			$self->{RANGE}->{'Value'} = [@arrofarr];
			# [[1a,1b,1c,1d],[2a,2b,2c,2d],[3a,3b,3c]];
		}
	}
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

