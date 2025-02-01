def calculateMaxSum(N,Arr):
    List=[]
    Sum=[]
    for i in range(N):
        for j in range(i,N+1):
            a=[]
            b=[]
            if i>0:
                b=Arr[:i]
            a=Arr[i:j]
            a.reverse()
            List=b+a+Arr[j:]
            c=0
            for k in range(len(List)):
                if k%2!=0:
                    c=c+List[k]
            Sum.append(c)
    return max(Sum)
            




def main():
    N = int(input())
    
    
    Arr = []
    for _ in range(N):
        Arr.append(int(input()))
        
    result = calculateMaxSum(N,Arr)
    
    
    print(result)
main()
