use File::Copy;
use strict;
use warnings;
use feature qw(say switch);
use Excel_lib;
use Win32::OLE qw(in with);
use Essent;
use Carp;
use Data::Dumper;

# TODO welche Funktionen brauche ich?
use Win32::OLE::Const 'Microsoft Excel';
# TODO Welche Konstanten sind mir dadurch zugänglich?
$Win32::OLE::Warn = 3;
# The value of $Win32::OLE::Warn determines what happens when an OLE error occurs.
# If it's 0, the error is ignored. If it's 2, or if it's 1 and the script is
# running under -w, the Win32::OLE module invokes Carp::carp(). If $Win32::OLE::Warn
# is set to 3, Carp::croak() is invoked and the program dies immediately.

my $Range = Range->new();
my $Command = Command->new();
my $Guess = Guess->new();

print "Modul excel_com.pm importiert.\n";

{
	package Option;
	$Option::VERSION = "0.1";
	
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
		return $self;
	}
}

{
	package Excelobject;
	$Excelobject::VERSION = "0.5";
	#@Excelobject::ISA = qw(Range);
	
	sub new {
		my $class = shift;
		my $self = {};
		#$self -> {EXCELFILE} = $_[0];
		bless($self, $class);
		#$Range = Range->new();
		
		## default-settings
		$self->{add_cell} = 0;	# add cell before data dumping
		$self->{transpose_level} = 0;	# insert formula instead of copy values
		$self->{confirm_execute} = 1;
		$self->{execute_show_all} = 0;
        $self->{check_exist} = 1;   # batch_col
        $self->{dest_in_cell} = 0;  # batch col
        $self->{execute_command} = 1;
		$self->{debug} = 0;
		
		## Command
		$Command->{add_cell} = 0;	# add cell before data dumping
		$Command->{transpose_level} = 0;	# insert formula instead of copy values
		$Command->{confirm_execute} = 1;
		$Command->{execute_show_all} = 0;
        $Command->{check_exist} = 1;   # batch_col
        $Command->{dest_in_cell} = 0;  # batch col
        $Command->{execute_command} = 1;
		$Command->{debug} = 0;
		
		## Guess
		$Guess->{add_cell} = 0;	# add cell before data dumping
		$Guess->{transpose_level} = 0;	# insert formula instead of copy values
		$Guess->{confirm_execute} = 1;
		$Guess->{execute_show_all} = 0;
        $Guess->{check_exist} = 1;   # batch_col
        $Guess->{dest_in_cell} = 0;  # batch col
        $Guess->{execute_command} = 1;
		$Guess->{debug} = 0;
		
		## Range
		$Range->{add_cell} = 0;	# add cell before data dumping
		$Range->{transpose_level} = 0;	# insert formula instead of copy values
		$Range->{confirm_execute} = 1;
		$Range->{execute_show_all} = 0;
        $Range->{check_exist} = 1;   # batch_col
        $Range->{dest_in_cell} = 0;  # batch col
        $Range->{execute_command} = 1;
		$Range->{debug} = 0;
		##
		
		return $self;
	}
	
    # TODO: aus Range heraus die hashes von self (Excelobject) erreichen
    sub get_option {
        my $self = shift;
        my $option = shift;
        my $val = $self->{$option};
        return $val;
    }
    
    # TODO
    sub option {
        my $self = shift;
        if (@_ % 2 != 0) {
            if (@_ > 1) {
                #croak "usage: tie \@array, $_[0], filename, [option => value]...";
                warn "usage: option [option => value]\n";
            } else {
                #return ($_[0] =~ /$regex/ && $_[0] !~ /^\s*\#/ ? 1 : 0);
                return ($self->{$_[0]} ? $self->{$_[0]} : warn "option $_[0] not recognized");
            }
            
        }
        my (%opts_in) = @_;
        my @valid_opts = qw(add_cell transpose_level confirm_execute execute_show_all check_exist dest_in_cell execute_command collect_execute debug);
        my $option;
        foreach my $opt_in (keys %opts_in) {
            warn "not recognized option: $opt_in" unless grep {$opt_in eq $_} @valid_opts;
            # settings anlegen $self->{confirm_execute} = 1 etc.
            $self->{$opt_in} = $opts_in{$opt_in};
			$Range->{$opt_in} = $opts_in{$opt_in};
			$Command->{$opt_in} = $opts_in{$opt_in};
			$Guess->{$opt_in} = $opts_in{$opt_in};
        }
        # settings dependencies, e.g. execute_command requires check_exist
        if ($self->get_option("execute_command")) {
            warn "options altered according to dependencies: check_exist => 1\n" unless $self->{check_exist};
            $self->{check_exist} = 1;
			$Range->{check_exist} = 1;
			$Command->{check_exist} = 1;
			$Guess->{check_exist} = 1;
        }
        if ($self->get_option("dest_in_cell")) {
            $self->{check_exist} = 1;
			$Range->{check_exist} = 1;
			$Command->{check_exist} = 1;
			$Guess->{check_exist} = 1;
        }
		print "";
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
		
		# TODO LASTLASTROW works!??
		#my $lastrow = $self->{WORKSHEET}->
		#nLastRow = .Cells.Find(What:="*", After:=.Cells(1), _
		#LookIn:=xlFormulas, LookAt:=xlWhole, _
		#SearchOrder:=xlByRows, _
		#SearchDirection:=xlPrevious).Row
		
		return $lastlast_row;
	}
	
	# TODO with OPTION & get_option!?
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
	
    sub cells_address {
        my $self = shift;
        # $E$1:$G$14 (range) || $E$1 (only one cell selected)
        my $range_address = $self->getrange;
        # parse address
        
        
        # @cells_address = ((1,2),(1,3),(2,2),(2,3));
        # TODO array of Cells-Objects!?
        my @cells_address = split /:/, $range_address;
        say "split erg ". scalar @cells_address;
        # default: more then one cell selected
        if (scalar @cells_address > 1) {
            $range_address =~ //;
        } else {
            # only one cell selected
            
            #inputFormula = "=SUM(R10C2:R15C2)"
            #MsgBox Application.ConvertFormula(formula:=inputFormula, fromReferenceStyle:=xlR1C1, toReferenceStyle:=xlA1)
            #my $conv = $self->{EXCEL}->ConvertFormula("ConvertFormula(formula:=$range_address, fromReferenceStyle:=xlR1C1, toReferenceStyle:=xlA1");
            #my $range_address = "\"=SUMME(B1:B3)\"";
            #my $str = "formula:$range_address,FromReferenceStyle:=xlA1,ToReferenceStyle:=xlR1C1";
            #say "str ". $str;
            #my $conv = $self->{EXCEL}->ConvertFormula($str);
            
            #$range_address = $self->A1toR1($range_address, 'string');
            #ActiveCell.Address(ReferenceStyle:=xlR1C1)
            say "org form ". $range_address;
            #say "conv ". $conv;
            @cells_address = ($range_address);
        }
        return @cells_address;
    }
    
	sub getrange {
		my $self = shift;
        ##############
		# 2
        # $E$1:$G$14 (range) || $E$1 (only one cell selected)
		#my $range_address = $self->{EXCEL}->Selection->Address();
        my $range_address = $self->{EXCEL}->Selection->Address("FromReferenceStyle:=xlA1,ToReferenceStyle:=xlR1C1");
		return $range_address;
        
#        #############
#		# 1
#        # $E$1
#		my $range_start = $self->{EXCEL}->Selection->Cells(1)->Address();
#        # 42
#        my $sel_count_cells = $self->count_selected_cells();
#        # $G$14
#		my $range_end = $self->{EXCEL}->Selection->Cells($sel_count_cells)->Address();
#        ##############

#        #############
#		# 1b, based on Cell-Objects
#        # $E$1
#		#my $range_start = $self->{EXCEL}->Selection->Cells(1)->Address();
#        my $Cells_start = $self->{EXCEL}->Selection->Cells(1);
#        # 42
#        $sel_count_cells = $self->{EXCEL}->Selection->Cells->Count();
#        # $G$14
#		my $Cells_end = $self->{EXCEL}->Selection->Cells($sel_count_cells);
#        ##############

        ##############
        # 3: return Range-Object
        #my $range_start = $self->{EXCEL}->Selection->Cells(1);
        #my $range_end = $self->{EXCEL}->Selection->Cells($sel_count_cells);
        #my $Range = $self->{WORKSHEET}->Range($range_start,$range_end);
        #return $Range;
        ##############
	}
    
    sub count_selected_cells {
        my $self = shift;
        my $sel_count_cells = $self->{EXCEL}->Selection->Cells->Count();
        return $sel_count_cells;
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
	## VER 1
    # col1                  # col2
    #i:\vera6 2012\def\TH01	i:\vera6 2012\def\TH02
    #COPY S01.tif debla.tif	COPY S01.tif debla.tif
    #COPY S02.tif intro1.tif	COPY S02.tif intro1.tif
    #COPY S03.tif intro2.tif	COPY S03.tif intro2.tif
    #COPY S04.tif stopp1.tif	COPY S04.tif stopp1.tif
    #COPY S05.tif VZ008_1.tif	COPY S05.tif VZ008_1.tif
    #COPY S06.tif VZ008_2.tif	COPY S06.tif VZ008_2.tif
    #COPY S07.tif VZ008_3.tif	COPY S07.tif VZ008_3.tif
    #COPY S08.tif VZ008_4.tif	COPY S08.tif VZ008_4.tif

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
            if ($self->get_option("dest_in_cell")) {
                push @collect_execute, [$self->batch_col2($row, $col)];
            } else {
                push @collect_execute, [$self->batch_col($row, $col)];
            }
            $col++;
			
		}
        die "ActiveCell is empty!" unless (@collect_execute);
		$self->option(collect_execute=>0);
        #$self->{collect_execute} = 0;
        
        # executes sammeln
        #my $Command = Command->new;
        foreach (@collect_execute) {
            my ($path_execute, $filename, $batch_string) = @$_;
            File::writefile($filename, $batch_string);
            $Command->execute_batch($path_execute, $filename) if $self->get_option("execute_command");
        }
	}
    ## batch_col_block_VER2
    # col1      # col2      # col3  # col4
    # copy		            copy	
    #f:\poly\TH\TH01\	.	f:\poly\TH\TH02	cache
    #S01.tif	debla.tif	S01.tif	debla.tif
    #S02.tif	intro1.tif	S02.tif	intro1.tif
    #S03.tif	intro2.tif	S03.tif	intro2.tif
    #S04.tif	stopp1.tif	S04.tif	stopp1.tif   
	sub batch_col_block_VER2 {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		if (!defined $row && !defined $col) {
			$self->activecell_pos();
			($row, $col) = @{$self->{activecell}{pos}};
		}
		
		my $jump_col;
        my @collect_execute;
        $self->option(collect_execute=>1);
		while (defined $self->{WORKSHEET}->Cells($row, $col)->{'Value'}) {
			## cell: operation-mode
			my $op = $self->get_op($row, $col);
			# one or two columns-jump? del -> 1, copy -> 2
			if ( ${$Command->{op}}{$op} eq "group1" ) {
				$jump_col = 1;
			} elsif ( ${$Command->{op}}{$op} eq "group2" ) {
				$jump_col = 2;
			}
			# $path_execute, $filename, $batch_string
            push @collect_execute, [$self->batch_col_VER2($row, $col, $op)];
            $col += $jump_col;
			
		}
        die "ActiveCell is empty!" unless (@collect_execute);
        $self->option(collect_execute=>0);
        
        # executes sammeln
        #my $Command = Command->new;
        foreach (@collect_execute) {
            my ($path_execute, $filename, $batch_string) = @$_;
            File::writefile($filename, $batch_string);
            $Command->execute_batch($path_execute, $filename) if $self->get_option("execute_command");
        }
	}
	
	sub get_op {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		if (!defined $row && !defined $col) {
			$self->activecell_pos();
			($row, $col) = @{$self->{activecell}{pos}};
		}
		my $op = Data::remove_ws $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
		#$self->option(op=>$op);
		warn "Command '$op' not valid?!" unless $Command->is_valid($op);
		return $op;
	}
	
    ## batch_col_VER2
    # col1      # col2   
    # copy		         
    #f:\poly\TH\TH01\	.
    #S01.tif	debla.tif
    #S02.tif	intro1.tif
    #S03.tif	intro2.tif
    #S04.tif	stopp1.tif
    sub batch_col_VER2 {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		if (!defined $row && !defined $col) {
			$self->activecell_pos();
			($row, $col) = @{$self->{activecell}{pos}};
		}
		my $op = shift || $self->get_op($row, $col);
        ## next cell: base/execute path
		$row++;
		my $path_source = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
        die "ActiveCell is empty!" unless ($path_source);
        die "'$path_source': $!" unless (-e $path_source) && $self->get_option("check_exist");
		$path_source .= "\\" unless $path_source =~ /\\$/;
        ##
		#my $filename = $path_source."excel_batch.bat";
		my $filename = $path_source."excel_".$op.".bat";
		## next cells
		my @array;
		if ( ${$Command->{op}}{$op} eq "group1" ) {
			@array = $self->batch_1col($row, $col, $op, $path_source);
		} elsif ( ${$Command->{op}}{$op} eq "group2" ) {
			@array = $self->batch_2col($row, $col, $op, $path_source);
		}
		
		my $batch_string = join "\n", @array;
        if ($self->get_option("collect_execute")) {
            return ($path_source, $filename, $batch_string);
        } else {
            File::writefile($filename, $batch_string);
            #my $Command = Command->new;
            $Command->execute_batch($path_source, $filename) if $self->get_option("execute_command");
        }
	}
	
	sub batch_1col {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $op = shift;
		my $path_source = shift;
        ## following cells
		$row++;
		my @source_file = $self->readcol($row, $col);
        my @array;
        foreach my $i (0 .. $#source_file) {
            my $source = $path_source.$source_file[$i];
            $source = '"'.$source.'"' if $source =~ /\s/;
            push @array, "$op $source";
        }
		return @array;
	}
	# TODO nicht von außen verwendbar machen! (-> eigenes Packes, not export!??)
	sub batch_2col {
		my $self = shift;
		my $row = shift;
		my $col = shift;
		my $op = shift;
		my $path_source = shift;
        ## next cell: destination path
        my $path_dest = $self->{WORKSHEET}->Cells($row, $col+1)->{'Value'};
        if ($path_dest) {
            #$path_dest =~ s/\\/\\\\/g;
            $path_dest .= "\\" unless $path_dest =~ /\\$/;
        } else {
            $path_dest = ".\\";
        }
        ##
        ## following cells
		$row++;
		my @source_file = $self->readcol($row, $col);
        my @dest_file = $self->readcol($row, $col+1);
        die "source- & dest-columns not equally long" if scalar @source_file != scalar @dest_file;
        my @array;
        foreach my $i (0 .. $#source_file) {
            my $source = $path_source.$source_file[$i];
            $source = '"'.$source.'"' if $source =~ /\s/;
            my $dest = $path_dest.$dest_file[$i];
            $dest = '"'.$dest.'"' if $dest =~ /\s/;
            if ($self->get_option("check_exist")) {
                die "source does not exist '$source'\n" unless -e $source;
                # nur einmal überprüfen!? Performance
                if ($path_dest =~ /^\w:\\/) {
                    # TODO mit Cwd?
                    # absolute path
                    # funktioniert nicht mit \\filer1\ etc
                    die "destination-path: $! '$path_dest' " unless -e $path_dest;
                } else {
                    # relative path
                    die "destination-path: $! '$path_source.$path_dest' " unless -e $path_source.$path_dest;
                }
            }
            push @array, "$op $source $dest";
        }
		return @array;
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
        ## first cell
		my $path_execute = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
        die "ActiveCell is empty!" unless ($path_execute);
        if ($self->get_option("check_exist")) {
            die "'$path_execute': $!" unless (-e $path_execute);
        }
		$row++;
		$path_execute =~ s/\\/\\\\/g;
		unless ($path_execute =~ /\\\\$/) {
			$path_execute .= "\\\\";
		}
		$filename = $path_execute."excel_batch.bat";
		my @array = $self->readcol($row, $col);
        if ($self->get_option("check_exist")) {
            #my $Guess = Guess->new();
            $Guess->{path_execute} = $path_execute;
            # surround source file with quotes (") if whitespace in filename/path
            @array = $Guess->parse(@array);
        }
		my $batch_string = join "\n", @array;
        if ($self->get_option("collect_execute")) {
            return ($path_execute, $filename, $batch_string);
        } else {
            File::writefile($filename, $batch_string);
            #my $Command = Command->new;
            $Command->execute_batch($path_execute, $filename) if $self->get_option("execute_command");
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
			if ($self->get_option("transpose_level")) {
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
		#my $Range = Range->new();
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
			if ($self->get_option("transpose_level")) {
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
		#my $Range = Range->new();
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
	
	# take argument as match || read value of active-cell (argument overrides active-cell)
	# goes down column from select row, remove if match value
	# stop if empty ("")/ not defined
	sub removerow_if {
		my $self = shift;
		my $match = shift;
		my $debug = $self->get_option('debug');
		$self->activecell_pos;
		my ($row, $col) = @{$self->{activecell}{pos}};
		unless ($match) {
			$self->activecell_val;
			$match = $self->{activecell}{value};
			unless ($match) {
				die "no parameter and no value in active cell to match against\n";
			}
			$row++;
		}
		my $count_rm_rows = 0;
		my $count_search_rows = 0;
		my $val = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
		while ( defined $val ) {
			last if $val eq '';
			if ($val eq $match) {
				$self->removerow($row);
				$count_rm_rows++;
			} else {
				$row++;
			}
			say "$row" if $debug;
			$count_search_rows++;
			$val = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
		}
		say "Rows searched for '$match': $count_search_rows - Rows removed: $count_rm_rows";
		warn "active cell is empty\n" if $count_search_rows == 0;
	}
    
    ## traverse_selection
    # higher-order perl idea
    # traverse/ loop through a Range of Cells
    # direction 1:
    # call traverse_selection through action-method (e.g. add 1 to current numeric value)
    # direction 2:
    # call action-method through traverse_selection
	
    sub traverse_selection {
        my $self = shift;
        # Range-Object
        # $self->{WORKSHEET}->Range
        my $Range = shift;
        #foreach my $Cell (@{$self->{WORKSHEET}->Range->Cells) {
        
        say "Range: $Range";
        my @keys = keys %{$Range};
        #say "Range: @keys";
        
        my $c = $Range->Count();
        say "c $c";
        #my $ad = $Range->Address();
        say "ad ".$Range->Address();
        #say "ad $ad";
        say "cells ".$Range->Cells;
        say "cuarr ".$Range->CurrentRegion;
        #my $cell_n = $Range->Next;
        #say "val ". $cell_n->{'Value'};
        say "rangerow ". $Range->Row;
        

        
        #say "@cells";
        #say "dd";
        #print Data::Dumper->Dumper($Range);
        
        say "----";
        
        #foreach my $Cell (@{$Range->Cells}) {
        #    $self->add1($Cell);
        #    # act.: $Cell->add1;
        #}
        
        
        
        
        #For Each c In Worksheets("Sheet1").Range("A1:D10").Cells
        #    If Abs(c.Value) < 0.01 Then c.Value = 0
        #Next
    }
    
    
    # print all values of worksheet (?)
    sub used_range {
        my $self = shift;
        my $everything=$self->{WORKSHEET}->UsedRange()->{Value};
        for (@$everything) {
            for (@$_) {
                print defined($_) ? "$_|" : "<undef>|";
            }
        }
    }
    
    sub add1 {
        my $self = shift;
        # $self->{WORKSHEET}->Cells($row, $col)
        my $Cell = shift;
        my $value = $Cell->{'Value'};
        $value++;
        #$self->{WORKSHEET}->Cells($self->{row},$self->{col})->{'Value'} = $insert;
        $Cell->{'Value'} = $value;
    }
    
    # Idee: traverse through selected cells
    # TODO Problem: convertFormula A1 to R1C1 funktioniert nicht! (in cells_adress)
    
    sub removerow_if2 {
		my $self = shift;
		my $match = shift;
		my $debug = $self->get_option('debug');
        # read cells start
		$self->activecell_pos;
		my ($row, $col) = @{$self->{activecell}{pos}};
		unless ($match) {
			$self->activecell_val;
			$match = $self->{activecell}{value};
			unless ($match) {
				die "no parameter and no value in active cell to match against\n";
			}
			$row++;
		}
		my $count_rm_rows = 0;
		my $count_search_rows = 0;
		my $val = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
        
        # array of cell-addresses
        my @cells_address = $self->cells_address();
        
        foreach my $cell (@cells_address) {
            
        }
        
        ####
		#while ( defined $val ) {
		#	last if $val eq '';
		#	if ($val eq $match) {
		#		$self->removerow($row);
		#		$count_rm_rows++;
		#	} else {
		#		$row++;
		#	}
		#	say "$row" if $debug;
		#	$count_search_rows++;
		#	$val = $self->{WORKSHEET}->Cells($row, $col)->{'Value'};
		#}
		#say "Rows searched for '$match': $count_search_rows - Rows removed: $count_rm_rows";
		#warn "active cell is empty\n" if $count_search_rows == 0;
	}
    
    sub removerow {
        my $self = shift;
        my $row = shift;
		#$self->{EXCEL}->Cells($row, $col)->Select;
		$self->{EXCEL}->Rows($row)->Select;
		$self->{EXCEL}->Selection->Delete;
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
    
    sub A1toR1 {
        my $self = shift;
        my $cells_object = shift;
        my $format = shift;
        my $row;
        my $col;
        if ($format eq 'string') {
            
        } else {
            
        }
        return ;
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
    @Guess::ISA = qw(Excelobject);
    
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
		return $self;
	}
	
    sub parse {
        my $self = shift;
        my @array = @_;
        my @parsed_array;
		# TODO alternative mit for oder foreach...
        for (my $i = 0; $i < scalar @array; $i++) {
			# copy || move etc
            if ( my ($op) = grep {$array[$i] =~ /^($_)/i} keys %{$self->{op}} ) {
                # TODO parse hier: valider op-mode?
				#
				# group1 || group2
                my $op_group = ${$self->{op}}{$op};
                #if ($self->{op}) {	# TODO ???
                #    $op_group = $self->{op};	#
                #    $array[$i] = $op_group." ".$array[$i];	# ???
				# $self->{GROUP1 || GROUP2}
				$array[$i] = $self->$op_group($array[$i]);
            }
        }
        return @array;
    }
    
    ## TODO group1: DEL MKDIR
#MKDIR
#i:\vera6 2012\def2\
#01
#02
#03
#04
#05
#06
#07
#08
#09
#10
#11
#12

    sub group1 {
        my $self = shift;
        my $string = shift;
		# only empty dirs
		#my @ops = qw(del mkdir rmdir);
		my @ops = @{${$Command->{group}}{group1}};
		my $regex_ops = join "|", @ops;
        $string =~ /($regex_ops)\s(\w.*)/i;
		my $op = $1;
		warn "operation not recognized '$string'\n" unless $op;
		$string = $2;
		warn "syntax not correct\n" unless $string;
		my $file1 = $self->rebuild($string);
		return "$op $file1";
    }
    
    ## group2: COPY MOVE REN
    sub group2 {
        my $self = shift;
        my $string = shift;
        # COPY 2012-03-15 Vereinsliste aktive.xlsx HIGH2 tee.txt
        my $path;
        #my @ops = qw(copy move ren);
		my @ops = @{${$Command->{group}}{group2}};
        my $regex_ops = join "|", @ops;
        #$string =~ /(copy|move|ren)\s(\w+.*)/i;
		$string =~ /($regex_ops)\s(\w+.*)/i;
        my $op = $1;
		warn "operation not recognized '$string'" unless $op;
        #$op = $self->{op} if $self->{op};
        $string = $2;
        my $files = $self->rebuild($string);
        return "$op $files";
    }
    
    sub rebuild {
        # '2012-03-15'
        # '2012-03-15 Vereinsliste'
        # '2012-03-15 Vereinsliste aktive.xlsx' -> exist!
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
            if (-e $self->{path_execute}.$string_rebuild) {
                $file_exist = 1;
                $string_rest = join " ", @string_split[$i..$#string_split];
                last;
            }
            $string_rebuild .= " ";
            $ws = 1;
        }
		# TODO funzt nicht mit 'COPY i:\vera6 2012\def\TH01\*.* i:\vera6 2012\def\_sammeln'
        die "first argument (file) does not exist!: ->$self->{path_execute}.$string_rebuild<-" unless $file_exist;
        die "second argument (destination) missing\n" unless $string_rest;
        $string_rest = $self->{path_dest}.$string_rest;
        $string_rest = '"'.$string_rest.'"' if $string_rest =~ /\s/;
        $string_rebuild = '"'.$string_rebuild.'"' if $ws;
        #return ($string_rebuild, $string_rest);
		return Data::remove_ws "$string_rebuild $string_rest";
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
        @{$self->{valid_ops}} = qw(copy move ren del mkdir);
		%{$self->{op}} = (copy=>"group2",
                           move=>"group2",
                           ren=>"group2",
                           del=>"group1",
						   mkdir=>"group1",
						   rmdir=>"group1");
        $self->{path_dest} = '';    # default for batch_col2
		# array: @{${$self->{group}}{group1}} & @{${$self->{group}}{group2}}
		# group1=>[del, mkdir];
		# group2=>[copy, move, ren];
		foreach my $op (keys %{$self->{op}}) {
			my $group = ${$self->{op}}{$op};
			push @{${$self->{group}}{$group}}, $op;
		}
		return $self;
	}
    sub is_valid {
        my $self = shift;
        my $op = shift;
        #return ($_[0] =~ /$regex/ && $_[0] !~ /^\s*\#/ ? 1 : 0);
        #my @g = grep {$op =~ /^$_$/i} (@{$self->{valid_ops}});
        #my $ret = ( @g ? 1 : 0 );
		my @ops = keys %{$self->{op}};
		return ( (grep {$op =~ /^$_$/i} @ops) ? 1 : 0 );
        #return ( (grep {$op =~ /^$_$/i} (@{$self->{valid_ops}})) ? 1 : 0 );
    }
    
    # bei sub in sub können Variablen nicht geteilt werden
	sub execute_batch {
		my $self = shift;
		my $path_execute = shift;
		my $filename = shift;
		if ($self->get_option("confirm_execute")) {
		#if ($self->{confirm_execute}) {
			print "Execute ", $filename, "?\n";
			if (Process::confirmJN()) {
				$self->execute($path_execute, $filename);
			} else {
				print "File not executed\n";
			}
		} else {
			$self->execute($path_execute, $filename);
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
			if ($self->get_option("execute_show_all")) {
				print "___EXECUTE LOG: $filename ___\n";
				print $result_execute;
				print "_________________\n";
				File::writefile($filename.".log", $result_execute);
			} else {	# show only errors
				my $first;
				### analyse result_lines; move to sub check_system_feedback
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
				###
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
	
	# TODO check_system_feedback, depending on operation
	sub check_system_feedback {
		my $self = shift;
		my $op = $self->get_option("op");
		
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
        #my $t = get_option("add_cell");
		if ($Range->get_option("add_cell")) {
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
}

{
	package Complete_Excel;
	@Complete_Excel::ISA = qw(Excelobject);
	
	sub new {
		my $class = shift;
		my $self = {};
		bless($self, $class);
        @{$self->{valid_ops}} = qw(copy move ren del);
		return $self;
	}
	
    ## getter, setter
    # valid_ops
    sub valid_ops {
        my $self = shift;
        if (@_) {
            push @{$self->{valid_ops}}, @_;
        } else {
            return @{$self->{valid_ops}};
        }
    }
    
    sub is_valid {
        my $self = shift;
        my $op = shift;
        #print 'this $string' =~ /[\$\@\%\*\&]+/ ? "yup($1)\n" . "nopen\n";
        #return ($_[0] =~ /$regex/ && $_[0] !~ /^\s*\#/ ? 1 : 0);
        my $ret = (grep {$op =~ /^$_$/i} @{$self->valid_ops} ? 1 : 0);
        return $ret;
        if ( grep {$op =~ /^$_$/i} @{$self->valid_ops} ) {
            return 1;
        } else {
            return 0;
        }
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
    
    ##

    $self->{WORKSHEET}->Shapes->AddPicture (
        $schnipp_abs,		# Filename As String
        1,					# LinkToFile As MsoTriState
        1,					# SaveWithDocument As MsoTriState
        490,				# Left As Single
        $top_align,	# Top As Single
        350,				# Width As Single
        40					# Height As Single
    );

1;

