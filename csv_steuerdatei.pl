use strict;
use warnings;
use Essent;
use feature qw/say switch/;
use Switch;

my $file = "z:\\Lernstand_VERA6 - 2013\\__VIPP\\Vera 6_Pilot_Steuerdatei_für Polyprint_final nach fächern.csv";
my $file_out = "z:\\Lernstand_VERA6 - 2013\\__VIPP\\vera6_2013\\vera6_2013_monster.dbf";

my @csv_file = File::readfile($file);
(my $field_ref, my $csv_matrix_ref) = Data::parse_csv(\@csv_file,";",1);
my %field = %$field_ref;

my @OUT;
my $mathe_heftnr = 1;
my $LfdID = 1;
foreach my $line_ref (@$csv_matrix_ref) {
    my @line_arr = @$line_ref;
    my $SchulID = $line_arr[$field{FachCode}].$line_arr[$field{BL_Code}].sprintf("%02d", $line_arr[$field{Schul_Code}]).$line_arr[$field{Schulart_Code}].$line_arr[$field{Klasse_Code}];
    my $Hefte = $line_arr[$field{Hefte}];
    my $th = $line_arr[$field{Testheft_Nr}];
    my $SchuelerID;
    my $TG_ENDE = "";
    my $CD;
    
    ## TL_Skript
    if ($line_arr[$field{FachCode}] == 1 || $line_arr[$field{FachCode}] == 3) {
        my $tl_skript_nr = tl_skript($line_arr[$field{FachCode}], $th);
        my @line_tl_skript = ($LfdID++,$SchulID,"",$line_arr[$field{FachName}],$line_arr[$field{Bundesland}],"",$tl_skript_nr,"",$line_arr[$field{Schulart_Code}],$line_arr[$field{Schulart}],$line_arr[$field{"Name der Schule"}],$line_arr[$field{Strasse}],$line_arr[$field{PLZ}],$line_arr[$field{Ort}],$TG_ENDE);
        push @OUT, join(";", @line_tl_skript) . "\n";
        
    }
    ##
    
    for (my $i=1; $i<= $Hefte; $i++) {
        ### SchuelerID-Zeilen
        ($CD, my @thvers) = th_ver($line_arr[$field{FachCode}], $th);
        my $th_modulus = ($i-1) % (scalar @thvers);
        my $heftnr;
        if ($line_arr[$field{FachCode}] == 3) {
            # Mathe: 1-40 durchlaufend
            $heftnr = $mathe_heftnr;
            if ($mathe_heftnr == 40) {
                $mathe_heftnr = 1;
            } else {
                $mathe_heftnr++;
            }
        } else {
            # Deutsch, Englisch: nur nach Vorgabe
            $heftnr = $thvers[$th_modulus];
        }
        my $Schueler = sprintf("%02d", $i) . sprintf("%02d",$heftnr);
        $SchuelerID = $SchulID . $Schueler;
        ###
        # LfdID;SchulID;SchuelerID;Fachname;Bundesland;Schuelernr;Heftnr;CD;Schulcode;Schulart;Schulname;SchulStrasse;SchulPLZ;Schulort;TG_ENDE
        my @line = ($LfdID++,$SchulID,$SchuelerID,$line_arr[$field{FachName}],$line_arr[$field{Bundesland}],$i,$heftnr,$CD,$line_arr[$field{Schulart_Code}],$line_arr[$field{Schulart}],$line_arr[$field{"Name der Schule"}],$line_arr[$field{Strasse}],$line_arr[$field{PLZ}],$line_arr[$field{Ort}],$TG_ENDE);
        # print "i: $i -- @line --\n";
        push @OUT, join(";", @line) . "\n";
    }
    
    ## X am Ende
    $TG_ENDE = "X";
    my @line_tg_ende = ($LfdID++,$SchulID,"",$line_arr[$field{FachName}],$line_arr[$field{Bundesland}],"","",$CD,$line_arr[$field{Schulart_Code}],$line_arr[$field{Schulart}],$line_arr[$field{"Name der Schule"}],$line_arr[$field{Strasse}],$line_arr[$field{PLZ}],$line_arr[$field{Ort}],$TG_ENDE);
    push @OUT, join(";", @line_tg_ende) . "\n";
    ##
    
}

## %%EOF
push @OUT, "%%EOF\n";
##
## head
unshift @OUT, "LfdID;SchulID;SchuelerID;Fachname;Bundesland;Schuelernr;Heftnr;CD;Schulcode;Schulart;Schulname;SchulStrasse;SchulPLZ;Schulort;TG_ENDE\n";
unshift @OUT, "(vera6_2013.dbm) STARTDBM\n";
unshift @OUT, "(vera6_2013.jdt) SETJDT\n";
unshift @OUT, "[(projects) (vera6_2013)] SETPROJECT\n";
unshift @OUT, "XGF\n";
unshift @OUT, "%!\n";
##

File::writefile($file_out,@OUT);
print "Fin!\n";

### fin!

### subs
sub th_ver {
    my $Fach = shift;
    my $th = shift;
    my @thvers;
    my $CD;
    switch ($Fach.$th) {
        ## Deutsch
        case "11+2" {@thvers = (1,2);$CD="ja"}
        case "13+4" {@thvers = (3,4);$CD="ja"}
        case "15+6" {@thvers = (5,6);$CD="ja"}
        case "17+8" {@thvers = (7,8);$CD="ja"}
        case "19+10" {@thvers = (9,10);$CD="ja"}
        case "111+12" {@thvers = (11,12);$CD="ja"}
        case "113+14" {@thvers = (13,14);$CD="ja"}
        ## Englisch
        case "21" {@thvers = (1);$CD="ja"}
        case "22" {@thvers = (2);$CD="ja"}
        case "23" {@thvers = (3);$CD="ja"}
        case "24" {@thvers = (4);$CD="ja"}
        case "25" {@thvers = (5);$CD="ja"}
        case "26" {@thvers = (6);$CD="ja"}
        case "27" {@thvers = (7);$CD="ja"}
        case "28" {@thvers = (8);$CD="ja"}
        ## Mathe
        case "31-40" {@thvers = (1 .. 40);$CD="nein"}
    }
    return ($CD,@thvers);
}

sub tl_skript {
    my $Fach = shift;
    my $th = shift;
    my $tl_skript_nr;
    switch ($Fach.$th) {
        ## Deutsch
        case "11+2" {$tl_skript_nr = 100}
        case "13+4" {$tl_skript_nr = 300}
        case "15+6" {$tl_skript_nr = 500}
        case "17+8" {$tl_skript_nr = 700}
        case "19+10" {$tl_skript_nr = 900}
        case "111+12" {$tl_skript_nr = 1100}
        case "113+14" {$tl_skript_nr = 1300}
        ## Mathe
        case "31-40" {$tl_skript_nr = 0}
    }
    return $tl_skript_nr;
}