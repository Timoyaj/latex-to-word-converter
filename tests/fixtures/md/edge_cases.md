# Edge Cases

## Currency Must Stay Prose

The price is $5.00 and the other is $10.00 today. A lone $ sign is fine.

Escaped dollars: the cost is \$5 and \$10 total.

## Code Fences Are Never Math

```python
price = 5  # costs $5 and $10
math = "$x^2$ not converted here"
```

Inline code `$x^2$` also stays literal.

## Math With Tricky Macros

$$
\left( x \middle| y \right) \quad \mathbb{R} \quad \text{J}\cdot\text{s}
$$

Operator form:

$$
\operatorname{argmax}_{x} f(x)
$$

## Bracket Display Math

\[
y = mx + b
\]

## Paren Inline Math

The slope \( m \) and intercept \( b \) determine the line.
