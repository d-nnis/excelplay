use feature "switch";
use File::Copy;
use strict;
use warnings;
# popri
#use lib('C:\\Dokumente und Einstellungen\\huesemann.POLYINTERN\\Eigene Dateien\\workspace\\V6\\');
# home
#use lib('C:\\Documents and Settings\\dennis\\My Documents\\workspace\\PMs\\');
use Essent;
# Excel-part
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;

print "Modul excel_com.pm importiert.\n";

{
	package Excelobject;
	
	sub new {
		my $class = shift;
		my $self = {};
		#$self -> {EXCELFILE} = $_[0];
		bless($self, $class);
		return $self;
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
	
	sub init_newExcelfile {
		my $self = shift;
		my $excelfile = shift;
		my $sheetno = shift;
		my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
			|| Win32::OLE->new('Excel.Application', 'Quit');
		my $Book;
		#print "excelfile:", $excelfile, "\n";
		if (-e $excelfile) {
			$Book = $Excel->Workbooks->Open($excelfile);
			$self->{WORKSHEET} = $Book->Worksheets($sheetno);
		} else {
			$Book = $Excel->Workbooks->Add;
			$Book->Worksheets($sheetno);
		}
	}
	
	sub bildd_dirs {
		my $self = shift;
		if (@_) {
			$self->{bildd_dir_rel} = $_[0] if ($_[0]);
			$self->{bildd_dir_nest} = $_[1] if ($_[1]);
		}
		return ($self->{bildd_dir_rel}, $self->{bildd_dir_nest});
	}
	
	sub last_row {
		my $self = shift;
		my $startrow = shift;
		my $col = shift;
		my $i = 0;
		while (defined($self->{WORKSHEET}->Cells($startrow + $i, $col)->{'Value'} )) {
			$i++;
		}
		return $i+$startrow-1;
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
	
	sub pos {
		my $self = shift;
		$self->{row} = shift;
		$self->{col} = shift;
	}
	
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
	
	sub incr_cell {
		my $self = shift;
		
	}
	
	sub add_head {
		my $self = shift;
		my @head = @_;
		$self->{var_sort} = \@head;
		# f�r insert_case
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
	
	# f�ge Werte von case in excel-sheet ein
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
	# erg�nze mit vals
	# f�ge sps-format von Importfi
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
				# f�ge in jede Zeile die Position von Importfi (bildd) ein.
				$self->add_excelval($row_itemnum, $col_itemnum+2, $self->{Importfi}) if (defined($self->{Importfi}));
				# f�ge in jede Zeile die Position von Fieldfil (Schnipsel) ein.
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
		#my $empty = $self->{WORKSHEET}->Cells(5,1)->{'Value'};
		#my $version = $self->{WORKSHEET}->Cells($self->{startrow}, $self->{startcolumn})->{'Value'};
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






1;
