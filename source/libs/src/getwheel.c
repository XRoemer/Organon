#include "getwheel.h"
#include <X11/Xlib.h>


int grab_wheel()
{
    Display *display;
    XEvent xevent;
    Window window;

    if( (display = XOpenDisplay(NULL)) == NULL ) {
        return -1;
    }

    window = DefaultRootWindow(display);
    XAllowEvents(display, AsyncBoth, CurrentTime);

    XGrabButton(display,
				 5,
				 AnyModifier,
				 window,
				 1,
				 ButtonPressMask,
				 GrabModeAsync,
				 GrabModeAsync,
				 None,
				 None
				 );

    XGrabButton(display,
				 4,
				 AnyModifier,
				 window,
				 1,
				 ButtonPressMask,
				 GrabModeAsync,
				 GrabModeAsync,
				 None,
				 None
				 );

    int ev = -1;
    int run = 1;

    while(run) {
        XNextEvent(display, &xevent);
        switch (xevent.type) {
            case ButtonPress:
                ev = xevent.xbutton.button;
                run = 0;
                break;
        }
    }

    XUngrabButton(display, 4, AnyModifier, window);
    XUngrabButton(display, 5, AnyModifier, window);

    XFlush(display);
    XCloseDisplay(display);

    return ev;

}

