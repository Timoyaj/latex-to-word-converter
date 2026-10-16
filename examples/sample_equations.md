# Mathematical Physics and Analysis Sample

This document provides a comprehensive test suite of mathematical expressions written in LaTeX notation embedded within Markdown. It includes inline Greek characters, fractions, summations, integrals, matrices, subscripts, superscripts, tables with formulas, and numbered display equations.

---

## 1. Introduction and Inline Notations

Let $\alpha, \beta, \gamma \in \mathbb{R}$ represent scalar parameters, while $\theta$ and $\phi$ denote spherical angular coordinates. In quantum mechanics, the wavevector is $\mathbf{k}$ with magnitude $k = \frac{2\pi}{\lambda}$, and the reduced Planck constant is $\hbar \approx 1.05457 \times 10^{-34} \text{ J}\cdot\text{s}$. 

Inline expressions such as $a_i x^i + b_j y^j = c_k z^k$ demonstrate mixed subscripting and superscripting. Standard kinetic energy is expressed as $T = \frac{1}{2} m v^2$, and the relativistic mass-energy equivalence is $E_0 = m_0 c^2$.

---

## 2. Fundamental Display Equations

Below are 10 display equations designed to evaluate OMML conversion, formatting, and automatic equation numbering.

### Equation 1: Quadratic Roots
The analytic solution for the roots of any quadratic polynomial $ax^2 + bx + c = 0$:

$$
x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}
$$

### Equation 2: Euler's Identity
Connecting fundamental mathematical constants $e$, $i$, $\pi$, $1$, and $0$:

$$
e^{i\pi} + 1 = 0
$$

### Equation 3: Normal Probability Distribution
The probability density function of a Gaussian distribution with mean $\mu$ and variance $\sigma^2$:

$$
f(x; \mu, \sigma^2) = \frac{1}{\sigma \sqrt{2\pi}} \exp\left( -\frac{(x - \mu)^2}{2\sigma^2} \right)
$$

### Equation 4: Gamma Function Integral
The integral representation of the Gamma function $\Gamma(z)$ for $\text{Re}(z) > 0$:

$$
\Gamma(z) = \int_0^\infty t^{z-1} e^{-t} \, dt
$$

### Equation 5: Riemann Zeta Function and Euler Product
The Dirichlet series and Euler product over all prime numbers $p$:

$$
\zeta(s) = \sum_{n=1}^{\infty} \frac{1}{n^s} = \prod_{p \text{ prime}} \frac{1}{1 - p^{-s}}
$$

### Equation 6: Continuous Fourier Transform
The continuous forward Fourier transform decomposing a function $f(x)$ into frequency components $\xi$:

$$
\hat{f}(\xi) = \int_{-\infty}^{\infty} f(x) e^{-2\pi i x \xi} \, dx
$$

### Equation 7: Maxwell's First Equation (Gauss's Law)
The differential formulation of Gauss's law for the electric field $\mathbf{E}$ given charge density $\rho$ and permittivity $\varepsilon_0$:

$$
\nabla \cdot \mathbf{E} = \frac{\rho}{\varepsilon_0}
$$

### Equation 8: Time-Dependent Schrödinger Equation
The fundamental evolution equation for a non-relativistic quantum state $|\Psi(t)\rangle$:

$$
i\hbar \frac{\partial}{\partial t}\Psi(\mathbf{r}, t) = \left[ -\frac{\hbar^2}{2m}\nabla^2 + V(\mathbf{r}, t) \right] \Psi(\mathbf{r}, t)
$$

### Equation 9: Multi-Order Taylor Series Expansion
The infinite series expansion of an infinitely differentiable function $f(x)$ centered around point $x_0$:

$$
f(x) = \sum_{k=0}^{\infty} \frac{f^{(k)}(x_0)}{k!} (x - x_0)^k
$$

### Equation 10: Matrix Eigenvalue System
The characteristic system for computing eigenvalues $\lambda$ and corresponding eigenvectors $\mathbf{v}$:

$$
\begin{pmatrix} a_{11} - \lambda & a_{12} \\ a_{21} & a_{22} - \lambda \end{pmatrix} \begin{pmatrix} v_1 \\ v_2 \end{pmatrix} = \begin{pmatrix} 0 \\ 0 \end{pmatrix}
$$

---

## 3. Mathematical Reference Table

The following table summarizes common physical quantities, symbols, and mathematical definitions:

| Quantity | Symbol | Defining Formula | Dimensions / Units |
| :--- | :---: | :--- | :--- |
| Angular Frequency | $\omega$ | $\omega = 2\pi f = \frac{2\pi}{T}$ | $\text{rad}\cdot\text{s}^{-1}$ |
| Momentum Operator | $\hat{p}_x$ | $\hat{p}_x = -i\hbar \frac{\partial}{\partial x}$ | $\text{kg}\cdot\text{m}\cdot\text{s}^{-1}$ |
| Capacitance | $C$ | $C = \frac{Q}{V} = \varepsilon \frac{A}{d}$ | $\text{Farad (F)}$ |
| Magnetic Flux | $\Phi_B$ | $\Phi_B = \iint_S \mathbf{B} \cdot d\mathbf{A}$ | $\text{Weber (Wb)}$ |
| Thermal Efficiency | $\eta$ | $\eta = 1 - \frac{T_C}{T_H} = \frac{W}{Q_H}$ | Dimensionless |
| Entropy Change | $\Delta S$ | $\Delta S = \int \frac{\delta Q_{\text{rev}}}{T}$ | $\text{J}\cdot\text{K}^{-1}$ |

---

## 4. Subscript and Superscript Combinations

Complex index notation frequently appears in tensor calculus and differential geometry:

- Covariant metric tensor: $g_{\mu\nu}$
- Contravariant metric tensor: $g^{\mu\nu}$
- Christoffel symbols of the second kind: $\Gamma^\lambda_{\mu\nu} = \frac{1}{2} g^{\lambda\sigma} \left( \partial_\mu g_{\nu\sigma} + \partial_\nu g_{\mu\sigma} - \partial_\sigma g_{\mu\nu} \right)$
- Riemann curvature tensor: $R^\rho_{\ \sigma\mu\nu} = \partial_\mu \Gamma^\rho_{\nu\sigma} - \partial_\nu \Gamma^\rho_{\mu\sigma} + \Gamma^\rho_{\mu\lambda}\Gamma^\lambda_{\nu\sigma} - \Gamma^\rho_{\nu\lambda}\Gamma^\lambda_{\mu\sigma}$
