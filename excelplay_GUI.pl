use strict;
use warnings;
use Prima qw(Application Buttons);
use excel_com;
my $ex = Excelobject->new;

## create window
my $window = Prima::MainWindow->new(
    text    => 'ExcelplayGUI',
    size    => [400,400]
);

$window->insert(
    Button=>
        #centered    => 1,
        text        => 'Init',
        pack    => {side => 'left', padx => 100 },
        #onClick     => sub {print "ACTION!\n"},
        onClick     => \&init,
    );
$window->insert(
    Button=>
        text    => 'Zeilen in eine Spalte',
        pack    => {side => 'left', padx => 10 },
        onClick => \&Zeilen_in_1Spalte,
    );
$window->insert(
    Button=>
        text    => 'Beenden',
        pack    => {side => 'left', pady => 10 },
        onClick  => sub { $::application-> close },
    );


run Prima;

## actions
sub init {
    $ex->init;
}

sub Zeilen_in_1Spalte {
    $ex->Zeilen_in_1Spalte;
}
