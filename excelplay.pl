use strict;
use warnings;
use excel_com;
use feature qw/say/;
use File::Basename;



#http://search.cpan.org/~gene/File-Backup-0.07/lib/File/Backup.pm

#chdir("f:\\Users\\d-nnis\\workspace\\excelplay") or die "Can't change directory: $!";

my $excelobj = Excelobject->new();

#$excelobj->init($excelfile,4);
$excelobj->init();

say "base, getrange";
say $excelobj->getrange();
#$excelobj->removerow_if2();


#$excelobj->option(debug=>1);
#$excelobj->removerow_if();
#$excelobj->removerow(4);
#say "range";
#print $excelobj->getrange();
#say "--";
#print "";
__END__

# TODO BUG001.xlsx: source does not exist '"i:\vera6 2012\int_THPage\VZ006_2"'

# TODO File::Basename
    #use File::Basename;
    #
    #($name,$path,$suffix) = fileparse($fullname,@suffixlist);
    #$name = fileparse($fullname,@suffixlist);
    #
    #$basename = basename($fullname,@suffixlist);
    #$dirname  = dirname($fullname);

## settings
#$excelobj->add_cell(0);
#$excelobj->transpose_level(0);
#$excelobj->{confirm_execute} = 0;

$excelobj->option(execute_command=>1, add_cell=>0, confirm_execute=>1, check_exist=>1);

# TODO confirm_execute kommt nicht an!
#$excelobj->option(confirm_execute=>1);
#$excelobj->option(execute_command=>1);

#$excelobj->{dest_in_cell} = 1;
# dest_in_cell requires check_exist
# $self->{execute_Command} requires check_exist
#$excelobj->batch_col();
# TODO wo write und execute???

#$excelobj->batch_col;
#$excelobj->batch_col_VER2;
## $excelobj->batch_col_block_VER2();
#$excelobj->batch_col_block();

my $F = ExForms->new();
#my @data = $F->job();


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

