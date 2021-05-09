from threading import Timer


def hello():
    print("hello, world")


class RepeatingTimer(Timer):
    def run(self):
        while not self.finished.is_set():
            self.function(*self.args, **self.kwargs)
            self.finished.wait(self.interval)


t = RepeatingTimer(5.0, hello)
t.start()
