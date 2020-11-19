#include <iostream>
#include <omp.h>
double step;
static long num_steps = 1000000000;
int main()
{
	int i;  double pi, sum, x;
	step = 1.0 / (double)num_steps;

	double start = omp_get_wtime();
	for (i = 0, sum = 0; i < num_steps; i++)
	{
		x = (i + 0.5)*step;
		sum += 4.0 / (1.0 + x*x);
	}
	double end = omp_get_wtime();
	pi = sum*step;
	printf("%.10f\n", pi);
	printf("time taken %f\n", end - start);
}