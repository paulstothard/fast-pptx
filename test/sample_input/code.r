for (column in names(genotypes)) {
  vcf %>%
    add_column(!!(column) := genotypes[[column]]) ->
    vcf
}
