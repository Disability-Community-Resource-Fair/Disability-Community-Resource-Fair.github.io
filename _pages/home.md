---
layout: default
permalink: /
title: Home
nav_order: 10
---

<div class="header-bar">
  <img src="{{ site.logo | prepend: '/assets/img/' | relative_url | bust_file_cache }}" class="mb-4" style="height:200px" alt="Logo is a yellow circle with a light blue ring around it containing the words Disability Community Resource Fair. The center shows a brown male hand giving an information document to a caucasian female hand wearing a red bracelet."/>
  <h1 style="font-size:4.5rem">{{ site.blog_name }}</h1>
  <h2>{{ site.blog_description }}</h2>
  <h2>
    <a href="https://www.facebook.com/profile.php?id=61553120680095&sfnsn=wa&mibextid=RUbZ1f">
      <i class="fa-brands fa-square-facebook"></i>
    </a> <b>&middot;</b>
    <a href="mailto:disabilityfair@gmail.com"><i class="fa-regular fa-envelope"></i></a>
  </h2>
</div>

<div class="vendor-buttons btn-toolbar justify-content-center my-2">
  <!-- <a href="/vendors" class="btn btn-primary">View Resources</a> -->
  <a href="/vendor-information" class="btn btn-secondary">Resource Vendor Registration</a>
  <!-- <a href="/sponsor-information" class="btn btn-info">Sponsor Registration</a> -->
</div>

<hr class="mt-0" />
<div class="post">
  <article>
    <div class="post-section">
      <h1 class="post-title text-center">
        Welcome!
      </h1>
      <p>The annual Disability Community Resource Fair is an event designed to help families discover and connect
        to local services in the area. The event will include family-friendly activities and a community resource fair
        featuring local service providers.</p>
      <h1 class="post-title text-center">Event Details</h1>
      <ul class="list-unstyled">
        <li><b>What:</b> Disability related resources and information</li>
        <li><b>When:</b> Saturday, May 2, 2026 from 10am to 1pm</li>
        <li><b>Where:</b> Mechanicsburg Middle School <a href="https://maps.google.com/?q=1750 S Market St, Mechanicsburg, PA 17055">1750 S Market St, Mechanicsburg, PA 17055</a></li>
        <li><b>Why:</b> It's a fun way to connect with your local community and learn about available resources!</li>
      </ul>
      <p>Still have questions? Want to volunteer? <a href="/contact-us">Contact us here</a></p>
    </div>
    <div class="post-section">
      <h1 id="Email" class="post-title text-center">
        Sign Up for Email Reminders
      </h1>
      <form
      action="https://formcarry.com/s/Y80V8S1AIqX"
      class="formcarryform"
      enctype="multipart/form-data">
        <div class="form-group">
          <label for="email_input">Email address</label>
          <input
            type="email"
            class="form-control"
            id="email_input"
            name="email"
            aria-describedby="emailHelp">
          <input type="hidden" name="purpose" value="Subscription to newsletter">
          <small id="emailHelp" class="form-text text-muted">Reminders include a few updates prior to the event and an invitaion to next year's event. We'll never share your email and you can unsubscribe at any time by replying to the email.</small>
        </div>
        <div class="text-center">
          <button type="submit" class="btn btn-primary mb-4 justify-content-center">Submit</button></div>
      </form>
    </div>
    
    <element class="sr-only">
      Hi screen-reader user! You found an easter egg! This text does not show up visually on the website. We are so
      glad you are here
    </element>
    {% include sponsor_thank.liquid %}
  </article>
</div>