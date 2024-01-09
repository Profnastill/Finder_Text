from multiprocessing import Process
import time

def function_1(arg1, arg2):
    # Ваш код функции 1
    time.sleep(arg1)
    print("Завеошение 1")
    return 1

def function_2(arg3, arg4):
    # Ваш код функции 2
    time.sleep(arg3)

    print("Завеошение 1")
    return 2

if __name__ == "__main__":
    arg1 = 1
    arg2 = 2
    arg3 = 3
    arg4 = 4

    p1 = Process(target=function_1, args=(arg1,arg2))
    p2 = Process(target=function_2, args=(arg3,arg4))

    p1.start()
    p2.start()

    p1.join()
    p2.join()
    print(p1)
    print(p2)