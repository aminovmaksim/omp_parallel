#include <iostream>
#include <omp.h>
double step;
static long num_steps = 1000000000;
int main()
{
	double pi = 0;
	double x = 0.0;
	double sum = 0.0;
	int i;
	step = 1.0 / (double)num_steps;

	double start = omp_get_wtime();
#pragma omp parallel for private(x) schedule(static) reduction(+:sum)
		for (i = 0; i < num_steps; i++) {
			x = (i + 0.5)*step;
			sum += 4.0 / (1.0 + x*x);
		}
	pi = sum * step;
	double end = omp_get_wtime();

	printf("%.10f\n", pi);
	printf("time taken %f\n", end - start);
}