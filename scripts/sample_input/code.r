#!/usr/bin/env Rscript
# Simulate some data
set.seed(42)
x <- 1:100
y <- x * 2 + rnorm(100, mean = 0, sd = 50)

# Create a linear model
model <- lm(y ~ x)

# Summary of the model
summary(model)

# Plot the data and the model
plot(x, y, main = "Linear Regression Model", xlab = "X", ylab = "Y")
abline(model, col = "red")

# Print the coefficients
print(coef(model))

