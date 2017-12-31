import numpy as np

def soft_threshold_singular_value(lambda1, W):
    U, s, V = np.linalg.svd(W, full_matrices=False)
    for i in range(s.shape[0]):
        v = s[i] - lambda1
        if v > 0.0:
            s[i] = v
        else:
            s[i] = 0.0
    return np.dot(U, np.dot(np.diag(s), V))

def hard_threshold_singular_value(max_rank, W):
    U, s, V = np.linalg.svd(W, full_matrices=False)
    for i in range(s.shape[0]):
        if i >= max_rank:
            s[i] = 0.0
    return np.dot(U, np.dot(np.diag(s), V))

def matrix_completion_common(A, P, approx_func, max_iter, err, B0):
    B_cur = B0
    P_rev = -1.0 * P + 1.0
    for i in range(max_iter):
        C = A * P + B_cur * P_rev
        B_succ = approx_func(C)
        if np.allclose(B_succ, B_cur, rtol=0.0, atol=err):
            return (True, i, B_succ)
        B_cur = B_succ
    return (False, max_iter-1, B_cur)

def matrix_completion_soft(A, P, lambda1, max_iter, err, B0):
    f = lambda W: soft_threshold_singular_value(lambda1, W)
    return matrix_completion_common(A, P, f, max_iter, err, B0)

def matrix_completion_hard(A, P, max_rank, max_iter, err, B0):
    f = lambda W: hard_threshold_singular_value(max_rank, W)
    return matrix_completion_common(A, P, f, max_iter, err, B0)

def test_matrix_completion():
    A = np.array([[1,2,0],[2,0,6],[0,6,9],[4,8,12]])
    P = np.array([[1,1,0],[1,0,1],[0,1,1],[0,1,0]])
    B0 = np.zeros(A.shape)
    lambda1 = 1e-1
    max_rank = 1
    max_iter = 10000
    err = 1e-9

    c,i,M = matrix_completion_soft(A, P, lambda1, max_iter, err, B0)
    print (c,i)
    print (M)
    c,i,M = matrix_completion_hard(A, P, max_rank, max_iter, err, B0)
    print (c,i)
    print (M)

test_matrix_completion()