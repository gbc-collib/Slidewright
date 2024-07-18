# Slidewright
A typescript library for manipulating existing powerpoint templates in memory. Including loading, and deleting


## Contributing
Feel Free to Contribute I will continue adding functionality as my use cases demand, but this project was only started because I could not find a solution using existing libraries
I did not start this project with it becoming a library in mind but I hope this code will save someone the headache I had to deal with


# Example
```ts
    const pe = new PowerPointEditor();
    await pe.loadPowerPoint('powerpoint.pptx')
    await pe.deleteSlide(1)
    await pe.savePowerPoint('out/powerpoint.pptx')
```


## TODO
- [ ] Delete Slide
- [ ] Replace text on a slide
- [ ] Replace Images
- [ ] Reorder Slides
