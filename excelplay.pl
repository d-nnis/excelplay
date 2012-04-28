use strict;
use warnings;
use excel_com;
use feature qw/say/;

#chdir("f:\\Users\\d-nnis\\workspace\\excelplay") or die "Can't change directory: $!";
my $excelobj = Excelobject->new();

$excelobj->init();
#$excelobj->init($excelfile,4);


## settings
#$excelobj->add_cell(0);
#$excelobj->transpose_level(0);
#$excelobj->{confirm_execute} = 0;

$excelobj->{confirm_execute} = 0;
#$excelobj->{dest_in_cell} = 1;
# dest_in_cell requires check_exist
# $self->{execute_Command} requires check_exist
#$excelobj->batch_col();
# TODO wo write und execute???
$excelobj->batch_col_VER2;
#$excelobj->batch_col_block();

# VER 1

# TODO VER 2: source und dest in getrenntem col
#copy	
#	
#f:\poly\TH\TH01\   .	
#S01.tif	debla.tif
#S02.tif	intro1.tif
#S03.tif	intro2.tif
#S04.tif	stopp1.tif

 
# TODO VER 3?
# : source und dest in einem col ??

# TODO work with selected area (Range)

#$excelobj->Zeilen_in_1Spalte(2,2,10,2);
# TODO add_cell-Problem bei Zeilen_in_1Spalte etc.
#$excelobj->add_cell(0);
#$excelobj->Zeilen_in_1Spalte();
#$excelobj->Spalten_in_1Zeile();

#$excelobj->active_cell('aim');

#print "transpose_level:",$excelobj->transpose_level,"\n";
#$excelobj->Zeilen_in_1Spalte(2,2,10,1);
# TODO
# regex, und regexp -handling vereinheitlichn

#$excelobj->regex(paste_resultaddcell);
#$excelobj->{regex} = '(\d)(\d)';

#$excelobj->regex('addcell');

# transpose_level 0: Formelbezug
# 1: Wert kopieren
#$excelobj->transpose_level(1);

#$excelobj->join_row_block(2,14);
#$excelobj->join_row_block();

#$excelobj->regex_col();



print "ende\n";
__END__

## merge cells
#$excelobj->join_row(2,14);
$excelobj->join_sep('-');
$excelobj->join_row_block(2,14);

##

#########
## alt ##
#my $colstart = 1;
#my $srow = 23;
#my $scol = 1;
#my @vals;
#while (my $vals_ref = [ $excelobj->readcol($colstart, 2) ] ) {
#	say $colstart.",";
#	my @vals = @{$vals_ref};
#	last if (scalar @vals) == 0;
#	foreach my $val (@vals) {
#		$excelobj->pos($srow, $scol);
#		$excelobj->writeval($val);
#		$srow++;
#	}
#	$colstart++;
#}




print "__End\n";