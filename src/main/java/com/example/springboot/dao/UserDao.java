package com.example.springboot.dao;

import com.example.springboot.entity.User;

import java.util.List;

public interface UserDao {

    List<User> findAll();

    User findById(int id);

    User save(User user);

    void deleteById(int id);
}
