# answer = 0

# for index in range(0, 11):
#     if index % 2 == 0:
#         print(index)
#         answer +=index

# print(n)

# sum([i for i in range(2, n + 1, 2)])


# print(sum([i for i in range(2, n + 1, 2)]))

# def example1():
#     for index in range(0, 11):
#         if index % 2 == 0:
#             print(index)
#             answer += index

#     print(n)

#     sum([i for i in range(2, n + 1, 2)])


# def example2():
#     numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

#     answer = 0

#     for index in range(0, len(numbers)):
#         answer = answer + numbers[index]

#     print(answer / len(numbers))


# import numpy as np
# def solution(numbers):
#     return np.mean(numbers)


# example2()


# def example3():
#     int answer = n / 10;
#     return (n * 12000) + ((k-answer == 0 ? 0 : k-answer) * 2000);

#     n = 10
#     k = 3

#     answer = n // 10

#     print((n * 12000) + (0 if k-answer == 0 else k-answer) * 2000)


# example3()


# def example4():
#     message = "happy birthday!"

#     print(len(message) * 2)


# example4()


# def example5():
#     num_list = [1, 2, 3, 4, 5]

#     answer = num_list[::-1]

#     print(answer)


# example5()


# def example6():  # 배열 자르기
#     # int[] answer = new int[num2-num1+1];

#     # for(int index=0;num1<=num2;num1++,index++) {
#     # 	answer[index] = numbers[num1];
#     # }

#     # return answer;

#     numbers = [1, 2, 3, 4, 5]
#     num1 = 1
#     num2 = 3

#     print(numbers[num1:num2+1])


# example6()


# def example7():
#     answer = 0
#     n = 15

#     if (n % 7 > 0):
#         answer = n // 7 + 1
#     else:
#         answer = n // 7

#     print(answer)


# example7()


# def example8():
#     numbers = [1, 2, 3, 4, 5]

#     for index in range(0,len(numbers)):
#         numbers[index] = numbers[index] *2

#     print(numbers)

# example8()


# def example9():

#     numbers = [0, 31, 24, 10, 1, 9]

#     numbers.sort()

#     print(numbers)


# example9()


def example10():
    my_string = "abcdef"
    letter = "f"

    str = my_string.replace("f", "")

    print(str)


example10()
