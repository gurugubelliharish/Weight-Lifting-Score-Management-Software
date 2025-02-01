import sys
def specialnumbers(Q,Queries):
    count=0
    List=[]
    for i in Queries:
        for A in range(2,i+1):
            for B in range(2,i+1):
                for C in range(2,i+1):
                    for D in range(2,i+1):
                        if (A**B)+(C**D)<=i:
                            List.append((A**B)+(C**D))
        for j in range(i+1):
            if j in List:
                count+=1
    return count
                
def main():
    Q = int(input())
    
    
    Queries = []
    for _ in range(Q):
        Queries.append(int(input()))
        
    result = specialnumbers(Q,Queries)
    
    
    print(result)
main()
