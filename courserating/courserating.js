
cs_elements.forEach(addCourseName)
cs_elements.forEach(setFncSubmit)

function addCourseName(value, index, array) {
    
    var tr = document.getElementById("courserating");
    var td = document.createElement("td");
    var form = document.createElement("form");
    form.setAttribute("id", "cs-form-" + value)
    var data = document.createElement("input");
    data.setAttribute("name", value);
    data.setAttribute("value", "1");
    data.setAttribute("type", "hidden");
    var input = document.createElement("input");
    input.setAttribute("name", "course");
    input.setAttribute("value", name);
    input.setAttribute("type", "hidden");
    var submit = document.createElement("input");
    submit.setAttribute("id", "submit_"+value);
    submit.setAttribute("type", "image");
    submit.setAttribute("src", "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAC8AAAApCAYAAAClfnCxAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAT/gAAE/4BB5Q5hAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAAeSURBVFiF7cExAQAAAMKg9U/tYwygAAAAAAAAAOAGHkUAAUS18VAAAAAASUVORK5CYII=");
    form.appendChild(input);
    form.appendChild(data);
    form.appendChild(submit);
    td.appendChild(form)
    tr.appendChild(td)
}

function setFncSubmit(value, index, array) {
        $(document).ready(function(){
            var form = $('form#cs-form-' + value),
              url = 'https://script.google.com/macros/s/AKfycbzq4-G18D19bUvdJaiv2ylvd5PACCBDajnOArzxQLazHDz0AlY/exec';
              form.submit(function(e){
                e.preventDefault();
              var jqxhr = $.ajax({
                url: url,
                method: "GET",
                dataType: "json",
                data: form.serialize()
              });
                $(".thanks").html("Thank you for getting in touch!").css("font-size","1rem");
                $(".form-control").remove();
                $("#courserating").remove();
                console.log(e.result)
              });
            });

}

