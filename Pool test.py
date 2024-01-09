from multiprocessing import Process,Pool
def func1(x):
    return x*x

def func2(y):
    return y**2



if __name__ == '__main__':
    pool = Pool()
    result1 = pool.apply_async(func1, (5,))
    result2 = pool.apply_async(func2, (6,))
    pool.close()
    print(result1.get())  # prints 25
    print(result2.get())  # prints 36