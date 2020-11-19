#include <stdio.h>
#include <stdlib.h>
#include <math.h>
#include <omp.h>

double f(double x)
{
    return sin(x);
}

int main()
{
    const long int N = 1000*1000*1000;
    double sum;
    double a = 0.0;
    double b = 1.0;
    double h = (b - a) / N;
    double time;
    int NumThreads;
#pragma omp parallel
    {
#pragma omp single
        {
            NumThreads = omp_get_num_threads();
            printf("Threads count is: %d\n", NumThreads);
        }
    }

    long int i;
    time = omp_get_wtime();
    sum = 0.0;
    for(i=0; i<N; i++)
    {
		sum += f(a + (i + 0.5)*h)*h;
    }
    time = omp_get_wtime() - time;
    printf("Serial time: %f\n", time);
    printf("Serial sum: %f\n", sum);
    
    time = omp_get_wtime();
    sum = 0.0;
	//TODO: why does it work incorrectly?
    #pragma omp parallel for num_threads(NumThreads)
    for(i=0; i<N; i++)
    {
		sum += f(a + (i + 0.5)*h)*h;
    }
    time = omp_get_wtime() - time;
    printf("Parallel time: %f\n", time);
    printf("Parallel sum: %f\n", sum);
    
    return 0;
}