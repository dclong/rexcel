library(XLConnect)
library(stringi)
library(diffr)

read_formula = function(file, sheet, cells = NULL, remove_space = FALSE, upper_case = FALSE) {
    wb = loadWorkbook(file)
    if (is.null(cells)) {
        nrow = getLastRow(wb, sheet = sheet)
        ncol = getLastColumn(wb, sheet = sheet)
        cells = cbind(rep(1:nrow, each = ncol), rep(1:ncol, times = nrow))
    }
    formula = data.frame(
        row = integer(0),
        col = integer(0),
        formula = character(0),
        stringsAsFactors = FALSE
    )
    for (i in 1:nrow(cells)) {
        row = cells[i, 1]
        col = cells[i, 2]
        r = try(getCellFormula(wb, sheet = sheet, row = row, col = col), silent = TRUE)
        if (class(r) != "try-error") {
            formula = rbind(formula, data.frame(row = row, col = col, formula = r, stringsAsFactors = FALSE))
        }
    }
    if (remove_space) {
        formula$formula = stri_replace_all_fixed(formula$formula, " ", "")
    }
    if (upper_case) {
        formula$formula = stri_trans_toupper(formula$formula)
    }
    formula
}

compare_formula = function(file1, sheet1, file2, sheet2, cells = NULL, upper_case = FALSE, remove_space = FALSE) {
    formula1 = read_formula(file1, sheet1, cells = cells, remove_space = remove_space, upper_case = upper_case)
    formula2 = read_formula(file2, sheet2, cells = cells, remove_space = remove_space, upper_case = upper_case)
    formula = merge(formula1, formula2, by = c("row", "col"), suffix = c("_1", "_2"), all = TRUE)
    formula$formula_1 = ifelse(is.na(formula$formula_1), '', formula$formula_1)
    formula$formula_2 = ifelse(is.na(formula$formula_2), '', formula$formula_2)
    formula[formula$formula_1 != formula$formula_2, ]
}

write_formula = function(formula) {
    file = tempfile()
    formula = stri_paste('[', formula[, 1], ', ', formula[, 2], ']: ', formula[, 3]) 
    writeLines(formula, con = file)
    file
}

visualize_formula = function(formula) {
    f1 = write_formula(formula[, c("row", "col", "formula_1")])
    f2 = write_formula(formula[, c("row", "col", "formula_2")])
    diffr(f1, f2)
}

f1 = "D:/Book1.xlsx"
f2 = "D:/Book2.xlsx"
formula = compare_formula(f1, 1, f2, 1, upper_case = TRUE, remove_space = TRUE)
formula = compare_formula(f1, 1, f2, 1, cells = formula[, 1:2])
visualize_formula(formula)
