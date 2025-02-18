from time import time

from ..config.logger_config import get_logger

# # logger = get_logger()

def timer_func(func):
    # This function shows the execution time of
    # the function object passed
    def wrap_func(*args, **kwargs):
        t1 = time()
        result = func(*args, **kwargs)
        t2 = time()
        logger.info(f'Function {func.__name__!r} executed in {(t2-t1):.4f}s')
        return result
    return wrap_func

def timer_func_v2(func):
    """
    Decorator function to measure the execution time of an asynchronous function.

    This decorator can be applied to an asynchronous function. It measures the time it takes for the
    decorated function to execute and logs the execution time.

    Parameters:
        func (callable): The asynchronous function to be decorated.

    Returns:
        callable: A wrapped version of the input function that measures and logs its execution time.

    Example:
        @timer_func_v2
        async def my_async_function():
            # Some asynchronous code
            pass

    Note:
        This decorator is designed for use with asynchronous functions and relies on Python's async/await syntax.

    """
    # This function shows the execution time of
    # the function object passed
    async def wrap_func(*args, **kwargs):
        t1 = time()
        result = await func(*args, **kwargs)
        t2 = time()
        logger.info(f'Function {func.__name__!r} executed in {(t2-t1):.4f}s')
        return result
    return wrap_func