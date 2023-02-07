#!/usr/bin/env perl
use strict;
use warnings;

my $cmd = '/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to csv';
chdir "data" or die;
foreach my $dir (glob "*") {
    chdir $dir or next;
    foreach my $xlsx (glob "*.xlsx") {
        my ($csv) = split m/\./, $xlsx;
        system "$cmd $xlsx";
        open (IN, "< ${csv}.csv") or die "${csv}.csv";
        my @in = <IN>;
        close(IN);
        open (OUT, "> ${csv}.csv") or die;
        foreach (@in) {
            my @data = split ",";
            next unless defined($data[2]);
            chomp($data[1], $data[2]);
            my ($num) = is_number($data[1]);
            $data[2] = lc($data[2]);
            print OUT "$data[2] $data[1]\n" if $num == 0;
        }
        close(OUT)
    }
    chdir ".." or die;
}
chdir '..' or die;
sub is_number {
    my $in = shift;
    my $reg1 = qr/^-?\d+(\.\d+)?$/;
    my $reg2 = qr/^-?0(\d+)?$/;
    my $out = 0;
    $out = 1 unless ($in =~ $reg1 && $in !~ $reg2);
    return $out;
    #是数字为0，不是数字为1
}