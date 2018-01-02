import numpy as np
from numpy.linalg import norm, svd
from scipy.sparse import coo_matrix, eye, spdiags
from scipy.sparse.linalg import cg

alpha = 0.5


def grad_psi(x):
    return np.sign(x) * (alpha / ((alpha * np.abs(x) + 1)**2) - alpha)


def linear_operator(x):
    col = np.where(~np.isnan(x.flatten()))[0]
    row = np.arange(col.size)
    data = np.ones(col.size)
    return coo_matrix((data, (row, col)), shape=(col.size, x.size))


def dp(x, beta=0.5, max_iter=200):
    A = linear_operator(x)
    b = x[~np.isnan(x)]
    x = np.nan_to_num(x)

    lambd = np.max(np.abs(A.T @ b))

    ds_best = np.inf
    lambd_best = lambd
    x_best = None

    for i in range(max_iter):
        lambd *= beta
        x = admm(x, A, b, lambd)
        ds = norm(A @ x.flatten() - b)

        if ds < ds_best:
            ds_best = ds
            lambd_best = lambd
            x_best = x

    return x_best


def admm(x, A, b, lambd, rho=1, max_iter=200, tol=1e-4):
    z = np.zeros(x.shape)
    I = eye(x.size)

    ata = A.T @ A
    atb = A.T @ b
    cga = ata + rho * I

    for i in range(max_iter):
        # Solve w
        tau = lambd * alpha / rho
        y = x + z / rho

        u, s, v = svd(y)
        w = u @ spdiags(np.maximum(s - tau, 0), 0, *x.shape) @ v

        # Quasi-Newton's method
        u, s, v = svd(x)
        dpsi = u @ spdiags(grad_psi(s), 0, *x.shape) @ v

        cgb = atb - ata @ x.flatten() - lambd * dpsi.flatten() \
        
            + rho * (w - x).flatten() - z.flatten()

        dx = cg(cga, cgb)[0].reshape(x.shape)

        err = norm(dx, 'fro') / norm(x, 'fro')

        if err < tol:
            x = x + dx
            break

        x += dx
        z -= rho * (w - x)

    return x


def example():
    n1 = 100
    n2 = 50
    r = 10
    x_real = np.dot(np.random.randn(n1, r), np.random.randn(r, n2))
    ratio = 0.8

    x_obs = x_real.copy()
    x_obs[np.random.rand(*x_real.shape) > ratio] = np.nan

    x_opt = dp(x_obs)
    print("relative error: {}".format(norm(x_opt - x_real) / norm(x_real)))


if __name__ == '__main__':
    example()