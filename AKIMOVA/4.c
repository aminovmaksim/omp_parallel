#include <stdio.h>
#include <omp.h>

int main()
{
	const int N = 100000;
	int i, j;
	int div, num_prime;
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

	num_prime = 0;
	time = omp_get_wtime();
	// TODO: try to accelerate the loop below
	for(i=2; i<=N; i++)
	{
		div = 0;
		for(j=2; j<=i; j++)
		{
			if(i % j == 0)
			{
				div++;
			}
		}
		if(div == 1)
		{
			#pragma omp atomic
			num_prime++;
		}
	}
	time = omp_get_wtime() - time;
	
	printf("Prime numbers: %d\n", num_prime);
	printf("Time: %f\n", time);
	return 0;
}