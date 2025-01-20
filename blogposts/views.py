from django.shortcuts import render
from django.views.generic import ListView

from .models import BlogPost


class BlogPost(ListView):
    model = BlogPost
    template_name = 'blogposts/blog.html'
    context_object_name = 'posts'
    queryset = BlogPost.objects.order_by('-date_added').all()



# def get_blogposts(request):
#     posts = Post.objects.all()
#     context = {'posts': posts}
#     return render(request, 'blogposts/blog.html', context=context)
