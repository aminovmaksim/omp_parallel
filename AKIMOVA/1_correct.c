#include <stdio.h>
#include <stdlib.h>
#include <math.h>
#include <omp.h>


int main()
{
    const long int N = 1000*1000*1000;
    double time;
    int NumThreads;
    
    double *a;
    a = (double*) malloc(N*sizeof(double));
    
    long int i;
    for(i=0; i<N; i++)
    {
		a[i] = (double) i;
    }
#pragma omp parallel
    {
#pragma omp single
        {
            NumThreads = omp_get_num_threads();
            printf("Threads count is: %d\n", NumThreads);
        }
    }
    time = omp_get_wtime();
    for(i=0; i<N; i++)
    {
		a[i] = sqrt(a[i]);
    }
    time = omp_get_wtime() - time;
    printf("Serial time: %f\n", time);
    
	//TODO: try to accelerate
    time = omp_get_wtime();
#pragma omp parallel for
    for(i=0; i<N; i++)
    {
		a[i] = sqrt(a[i]);
    }
    time = omp_get_wtime() - time;
    printf("Parallel time: %f\n", time);
    printf("a[1]=%f\n",a[1]);
    free(a);
    return 0;
}