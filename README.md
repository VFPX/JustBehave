# JustBehave

**Extends the behavior of any Visual FoxPro baseclass or custom framework class without the need for additional code.**

If a given class requires a behavior which is not inherited in a particular class hierarchy branch, the particular code will be "cut and pasted" into the individual control or an entirely new class branch.

Behavior code is loaded once for the entire application and that all registered controls share the common behavior. The overall result is a smaller footprint and a potentially more stable operating environment.

The goals of this project are:

* Allow custom behaviors for base or thin  UI classes in a centralized manner which will allow multiple behaviors to be added to produce rich UI interfaces with virtually NO programming.  

* Improve UI performance and stability by centralizing behavior code and allowing lightweight controls.

## Try it

    Do form demo1

This form is a simple demonstration of how a behavior class interacts with Base VFP classes.

The Behaviors framework utilizes an expandable set of classes which are wrapped by session class observers. All the housekeeping and plumbing is handled by the wrapper class and the only method in the behavior class which normally requires development is the "Handle" method. Each behavior class identifies which events it will bind to in a custom property "cEvents", further the event binding can be limited to only those objects which have a required property or method as specified in "cReuiquiredProperties" and "cRequiredMethods".

When executing the "handle" method, the behavior object will "know" the calling object reference, the form object reference, and the method which invoked the behavior in:

    Behavior::oCaller
    Behavior::oForm
    Behavior:cEvent

Because the behavior object is global, data session swapping may become necessary for some behaviors. This is accomplished with the Behavior::lSwapDataSession property. If set to True the data session will automatically follow the data session of the calling object.

The wrapper class definition is contained in the behaviorfunctions.prg. The behaviorfunctions.prg also supplies the registration process which drills through containers looking for objects to associate and bind to behaviors. A form would only need a few lines of code to enable this technique.

    Form::Load
    Set Procedure To programs\BehaviorFunctions.prg
    Set Classlib To libraries\Behaviors.vcx

    Form::Init
    subscribe(This)

Base class controls subscribe to behaviors by naming the behavior class in its Tag property.

If more that one behavior is required the additional behaviors are added to the Tag pproperty with space delimiters.  

    Testbox::Tag
    cbackgroundhighlighter cmouseenterselect

If the requested behavior is not found it is ignored, this can be improved to log the error so that testing would reveal typos and such.

Currently there is VERY little error checking and virtually no error handling. This is more of a proof of concept exercise than a production tool. As time permits I will be adding more useful behaviors and add more strenuous error handling.