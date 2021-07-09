package com.example.springboot.service;

import com.example.springboot.entity.User;

import java.util.List;

public interface UserService {

    List<User> findAll();

    User findById(int id);

    User save(User user);

    void deleteById(int id);
}
