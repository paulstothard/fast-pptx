#!/usr/bin/perl

use strict;
use warnings;

# Check for input file argument
my $filename = $ARGV[0] or die "Usage: $0 filename\n";

# Initialize counters
my ($line_count, $word_count, $char_count) = (0, 0, 0);

# Open the file for reading
open my $fh, '<', $filename or die "Could not open '$filename' $!\n";

# Process the file line by line
while (my $line = <$fh>) {
    $line_count++;
    $char_count += length($line);
    $word_count += scalar split(/\s+/, $line) if $line =~ /\S/;
}

# Close the file
close $fh;

# Print the results
print "Lines: $line_count\nWords: $word_count\nCharacters: $char_count\n";
